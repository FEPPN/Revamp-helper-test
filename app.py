#!/usr/bin/env python3
"""The copywriter-facing tool: a simple web page (built with Streamlit) that
wraps every step of the revamp report pipeline behind one form and one
button. No Antigravity/Claude Code needed to use this — just this page.

Local run (for testing):     streamlit run app.py
Hosted run (Streamlit Cloud): this file is the entry point Streamlit Cloud
                               runs automatically once the repo is connected.

Where the API keys live: Streamlit "Secrets" (Settings > Secrets on
Streamlit Cloud, or a local .streamlit/secrets.toml for testing). Never in
this file, never in the repo. See secrets.toml.example for the shape.
"""

import io
import json
import subprocess
import sys
import tempfile
from datetime import date, timedelta
from pathlib import Path

import requests
import streamlit as st
from bs4 import BeautifulSoup
from googleapiclient.discovery import build as build_google_service
from openpyxl import Workbook

from scripts.build_report import (
    build_aide_sheet, build_serp_sheet, build_ahrefs_sheet, build_gsc_sheet, build_competitor_sheet,
    build_entity_sheet,
)
from scripts.fetch_gsc import get_credentials_from_values, run_query as gsc_run_query
from scripts.fetch_serp import fetch_serp, build_serp_json
from scripts.find_pages import search_site, domain_of, TARGET_DOMAIN, COMPETITOR_DOMAINS
from scripts.scrape_competitors import extract_structure, HEADERS as SCRAPE_HEADERS
from scripts.parse_ahrefs_csv import (
    find_column, to_number, parse_intents, FIELD_VARIANTS, INTENT_VARIANTS, COMBINED_INTENTS_VARIANTS,
)
import csv as csv_module

GSC_SITE_PROPERTY = "https://www.papernest.com/"
ENTITY_SCRIPT = Path(__file__).resolve().parent / "scripts" / "analyze_entities.py"

st.set_page_config(page_title="Revamp Report — papernest FR", page_icon="📊")
st.title("📊 Revamp Report — génération automatique")
st.caption("Mot-clé + marque → rapport Excel (SERP, Ahrefs, GSC, Concurrents). "
           "Aucune installation, aucune clé à toi de gérer.")


# ---------------------------------------------------------------------------
# Step 1 — target page + competitor pages (SerpAPI `site:` search)
# ---------------------------------------------------------------------------
def step_find_pages(keyword, serpapi_key):
    result = {"target": None, "target_alternatives": [], "competitors": []}

    target_results = search_site(keyword, TARGET_DOMAIN, serpapi_key)
    if target_results:
        result["target"] = {"url": target_results[0]["link"], "title": target_results[0].get("title", "")}
        for item in target_results[1:]:
            if TARGET_DOMAIN in domain_of(item["link"]):
                result["target_alternatives"].append({"url": item["link"], "position": item.get("position")})

    for domain in COMPETITOR_DOMAINS:
        comp_results = search_site(keyword, domain, serpapi_key)
        if comp_results:
            result["competitors"].append({
                "site": domain, "url": comp_results[0]["link"],
                "title": comp_results[0].get("title", ""), "no_dedicated_avis_page": False,
            })
        else:
            result["competitors"].append({"site": domain, "url": None, "no_dedicated_avis_page": True})

    return result


# ---------------------------------------------------------------------------
# Step 2 — competitor H1/H2/H3 scraping
# ---------------------------------------------------------------------------
def step_scrape_competitors(pages):
    results = []
    for comp in pages["competitors"]:
        if comp.get("no_dedicated_avis_page") or not comp.get("url"):
            results.append({
                "site": comp["site"], "url": comp.get("url") or "", "h1": "(aucune page dédiée trouvée)",
                "structure": [], "summary": "Aucune page dédiée trouvée pour ce mot-clé sur ce site.",
                "no_dedicated_avis_page": True,
            })
            continue
        try:
            r = requests.get(comp["url"], headers=SCRAPE_HEADERS, timeout=20)
            r.raise_for_status()
            h1, summary, structure = extract_structure(r.text)
            results.append({"site": comp["site"], "url": comp["url"], "h1": h1, "structure": structure, "summary": summary})
        except requests.RequestException as e:
            results.append({
                "site": comp["site"], "url": comp["url"], "h1": "(échec du téléchargement)",
                "structure": [], "summary": f"Erreur : {e}", "no_dedicated_avis_page": True,
            })
    return results


# ---------------------------------------------------------------------------
# Step — entity gap (page cible vs concurrents), via Google Natural Language
# API. Runs the pack's stdlib-only analyze_entities.py as a subprocess (same
# proven integration pattern as Page Audit) — no reimplementation.
# ---------------------------------------------------------------------------
def step_entity_gap(page_url, competitor_urls, api_key):
    if not competitor_urls:
        return None, "Aucune page concurrente valide à comparer."
    env = {**__import__("os").environ, "GOOGLE_NLP_API_KEY": api_key}
    cmd = [sys.executable, str(ENTITY_SCRIPT), page_url, "--language", "fr",
           "--gap-vs", ",".join(competitor_urls), "--json"]
    try:
        proc = subprocess.run(cmd, capture_output=True, text=True, timeout=90, env=env)
    except subprocess.TimeoutExpired:
        return None, "Analyse d'entités trop longue (timeout)."
    if proc.returncode != 0:
        return None, proc.stderr.strip() or "Échec de l'analyse d'entités."
    try:
        data = json.loads(proc.stdout)
    except json.JSONDecodeError:
        return None, "Réponse d'analyse d'entités illisible."
    return data.get("gap"), None
    return results


# ---------------------------------------------------------------------------
# Step 3 — Ahrefs: CSV upload (always available) — API key path can be added
# later the same way once workspace permissions allow creating one.
# ---------------------------------------------------------------------------
def step_parse_ahrefs_csv(uploaded_file):
    text = uploaded_file.getvalue().decode("utf-8-sig")
    reader = csv_module.DictReader(io.StringIO(text))
    fieldnames_lower = {name.strip().lower(): name for name in reader.fieldnames}
    col = {key: find_column(fieldnames_lower, variants) for key, variants in FIELD_VARIANTS.items()}
    if not col["keyword"]:
        raise ValueError(f"Colonne 'Keyword' introuvable. Colonnes trouvées : {reader.fieldnames}")
    intent_cols = {key: find_column(fieldnames_lower, variants) for key, variants in INTENT_VARIANTS.items()}
    combined_intents_col = find_column(fieldnames_lower, COMBINED_INTENTS_VARIANTS)

    rows = []
    for raw_row in reader:
        keyword = (raw_row.get(col["keyword"]) or "").strip()
        if not keyword:
            continue
        cpc = to_number(raw_row.get(col["cpc"])) if col["cpc"] else None
        intents = parse_intents(raw_row, intent_cols, combined_intents_col)
        rows.append({
            "keyword": keyword,
            "volume": to_number(raw_row.get(col["volume"])) if col["volume"] else None,
            "difficulty": to_number(raw_row.get(col["difficulty"])) if col["difficulty"] else None,
            "cpc": cpc,
            "traffic_potential": to_number(raw_row.get(col["traffic_potential"])) if col["traffic_potential"] else None,
            "intents": intents,
            "parent_topic": (raw_row.get(col["parent_topic"]) or "").strip() if col["parent_topic"] else None,
        })
    return rows


# ---------------------------------------------------------------------------
# Step 4 — GSC: automatic via OAuth (no CSV, no manual step). Credentials
# come from a one-time local browser login (see scripts/fetch_gsc.py); the
# hosted app only ever refreshes the stored refresh_token, never logs in
# interactively itself.
# ---------------------------------------------------------------------------
def step_fetch_gsc(target_url, client_id, client_secret, refresh_token):
    creds = get_credentials_from_values(client_id, client_secret, refresh_token)
    service = build_google_service("searchconsole", "v1", credentials=creds)

    end = date.today()
    start = end - timedelta(days=180)
    raw_rows = gsc_run_query(
        service, GSC_SITE_PROPERTY, start.isoformat(), end.isoformat(),
        dimensions=["query"], row_limit=1000, page_filter=target_url,
    )

    rows = []
    for r in raw_rows:
        rows.append({
            "query": r["keys"][0],
            "clicks": r.get("clicks", 0),
            "impressions": r.get("impressions", 0),
            "ctr": r.get("ctr", 0.0),
            "position": r.get("position", 0.0),
        })
    return rows


# ---------------------------------------------------------------------------
# Assembly — same logic as build_report.py, fed from in-memory data instead
# of files (the CLI script still works standalone for local/manual runs).
# ---------------------------------------------------------------------------
def assemble_workbook(keyword, brand, serp_data, ahrefs_rows, gsc_rows, competitors_data, page_url,
                       entity_gap=None):
    wb = Workbook()
    build_aide_sheet(wb)

    with tempfile.TemporaryDirectory() as tmp:
        tmp = Path(tmp)
        serp_path = tmp / "serp.json"
        serp_path.write_text(json.dumps(serp_data, ensure_ascii=False), encoding="utf-8")
        build_serp_sheet(wb, serp_path)

        ahrefs_path = tmp / "ahrefs.json"
        ahrefs_path.write_text(json.dumps(ahrefs_rows, ensure_ascii=False), encoding="utf-8")
        build_ahrefs_sheet(wb, ahrefs_path, brand)

        gsc_path = tmp / "gsc.csv"
        with open(gsc_path, "w", newline="", encoding="utf-8") as f:
            writer = csv_module.writer(f)
            writer.writerow(["query", "clicks", "impressions", "ctr", "position"])
            for r in gsc_rows:
                writer.writerow([r["query"], r["clicks"], r["impressions"], r["ctr"], r["position"]])
        build_gsc_sheet(wb, gsc_path, page_url)

        comp_path = tmp / "competitors.json"
        comp_path.write_text(json.dumps(competitors_data, ensure_ascii=False), encoding="utf-8")
        build_competitor_sheet(wb, comp_path)

        if entity_gap:
            build_entity_sheet(wb, entity_gap, page_url)

    buffer = io.BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    return buffer


# ---------------------------------------------------------------------------
# The form
# ---------------------------------------------------------------------------
with st.form("revamp_form"):
    keyword = st.text_input("Mot-clé (FR)", placeholder="ex: avis edf")
    brand = st.text_input("Marque / fournisseur", placeholder="ex: EDF")
    ahrefs_csv = st.file_uploader(
        "Export CSV Ahrefs (Keywords Explorer → ton mot-clé → Matching terms → Terms match: All → Export)",
        type=["csv"],
    )
    submitted = st.form_submit_button("Générer le rapport")

if submitted:
    if not keyword or not brand:
        st.error("Mot-clé et marque sont obligatoires.")
        st.stop()
    if not ahrefs_csv:
        st.error("Dépose le fichier CSV Ahrefs — sans lui, l'onglet mots-clés secondaires ne peut pas être rempli.")
        st.stop()

    serpapi_key = st.secrets.get("SERPAPI_API_KEY")
    if not serpapi_key:
        st.error("Clé SerpAPI absente des Secrets de l'app — contacte l'admin de l'outil.")
        st.stop()

    gsc_client_id = st.secrets.get("GSC_CLIENT_ID")
    gsc_client_secret = st.secrets.get("GSC_CLIENT_SECRET")
    gsc_refresh_token = st.secrets.get("GSC_REFRESH_TOKEN")
    google_nlp_key = st.secrets.get("GOOGLE_NLP_API_KEY")

    with st.status("Génération en cours...", expanded=True) as status:
        st.write("🔎 Recherche de la page cible et des pages concurrentes...")
        pages = step_find_pages(keyword, serpapi_key)
        if not pages["target"]:
            status.update(label="Échec", state="error")
            st.error("Aucune page papernest.com trouvée pour ce mot-clé. Vérifie l'orthographe.")
            st.stop()
        st.write(f"→ Page cible : {pages['target']['url']}")
        if pages["target_alternatives"]:
            st.warning(f"⚠️ Cannibalisation possible — autres pages papernest.com trouvées : "
                       f"{[a['url'] for a in pages['target_alternatives']]}")

        st.write("📄 Lecture des pages concurrentes...")
        competitors_data = step_scrape_competitors(pages)

        st.write("🔍 Récupération des données SERP Google...")
        raw_serp = fetch_serp(keyword, serpapi_key)
        serp_data = build_serp_json(raw_serp, keyword)

        st.write("📈 Récupération des mots-clés Ahrefs...")
        ahrefs_rows = step_parse_ahrefs_csv(ahrefs_csv)
        st.write(f"→ {len(ahrefs_rows)} mots-clés secondaires trouvés")

        st.write("📊 Récupération des stats Google Search Console...")
        if gsc_client_id and gsc_client_secret and gsc_refresh_token:
            try:
                gsc_rows = step_fetch_gsc(pages["target"]["url"], gsc_client_id, gsc_client_secret, gsc_refresh_token)
                st.write(f"→ {len(gsc_rows)} requêtes trouvées")
            except Exception as e:
                st.warning(f"Récupération Search Console impossible ({e}) — le rapport sera généré sans cet onglet rempli.")
                gsc_rows = []
        else:
            st.warning("Identifiants Search Console absents des Secrets de l'app — onglet GSC non rempli.")
            gsc_rows = []

        st.write("🧩 Analyse de l'écart d'entités vs concurrents...")
        entity_gap = None
        if google_nlp_key:
            comp_urls = [c["url"] for c in pages["competitors"] if c.get("url")]
            entity_gap, entity_err = step_entity_gap(pages["target"]["url"], comp_urls, google_nlp_key)
            if entity_err:
                st.warning(f"Analyse d'entités indisponible ({entity_err}) — onglet Entités non rempli.")
        else:
            st.warning("Clé Google NLP absente des Secrets de l'app — onglet Entités non rempli.")

        st.write("📁 Assemblage du fichier Excel...")
        buffer = assemble_workbook(keyword, brand, serp_data, ahrefs_rows, gsc_rows, competitors_data,
                                    pages["target"]["url"], entity_gap=entity_gap)

        status.update(label="Terminé ✅", state="complete")

    st.success("Rapport généré !")
    safe_keyword = keyword.replace(" ", "_")
    st.download_button(
        "⬇️ Télécharger le rapport Excel",
        data=buffer,
        file_name=f"{safe_keyword}_FR_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
