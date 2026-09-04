#!/usr/bin/env python3
"""Fetches live Google SERP data (FR) via SerpAPI and writes it in the exact
JSON shape build_report.py's SERP tab expects. Fully automatic, no LLM
involved — the per-result "note" is SerpAPI's own snippet, not an
AI-written summary.

Usage:
  python fetch_serp.py --keyword "avis edf" --out serp.json
"""

import argparse
import json
import os
import re
import sys

import requests

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

SERPAPI_URL = "https://serpapi.com/search.json"


def fetch_serp(keyword, api_key):
    """No free equivalent exists for this call - unlike find_pages.py's
    simple "top URL for a site:" search, this pulls Google-SERP-specific
    data (People Also Ask, related searches, AI Overview presence,
    Knowledge Graph presence) that a free scraper/DuckDuckGo cannot
    replicate with the same fidelity. So instead of a fallback, this
    degrades gracefully: on any SerpAPI failure (quota exhausted, invalid
    key, etc.) it returns {"_serp_error": <message>} instead of raising -
    the caller/report must then say "données SERP indisponibles"
    explicitly, never silently show empty PAA/related searches as if that
    were a real finding (no PAA questions exist for this keyword)."""
    params = {
        "engine": "google",
        "q": keyword,
        "gl": "fr",
        "hl": "fr",
        "api_key": api_key,
    }
    try:
        r = requests.get(SERPAPI_URL, params=params, timeout=30)
        r.raise_for_status()
        return r.json()
    except requests.exceptions.RequestException as e:
        # requests' own exception message embeds the full request URL,
        # api_key included - this message ends up in the Google Sheet report
        # (write_serp_sheet), so the key must never reach it in clear text.
        safe_message = re.sub(r"api_key=[^&\s]+", "api_key=***", str(e))
        return {"_serp_error": safe_message}


def build_serp_json(raw, keyword):
    if raw.get("_serp_error"):
        return {
            "keyword": keyword,
            "market": "FR (google.fr)",
            "organic": [],
            "paa": [],
            "related_searches": [],
            "ai_overview_present": None,
            "ai_overview_summary": "",
            "knowledge_graph_present": None,
            "serp_unavailable": f"Données SERP indisponibles (SerpAPI a échoué : {raw['_serp_error']}) "
                                 f"— champs ci-dessus non renseignés, pas 'aucun résultat trouvé'.",
        }

    organic = []
    for item in raw.get("organic_results", [])[:10]:
        organic.append({
            "position": item.get("position"),
            "title": item.get("title", ""),
            "link": item.get("link", ""),
            "note": item.get("snippet", ""),
        })

    paa = [q.get("question", "") for q in raw.get("related_questions", []) if q.get("question")]

    related_searches = [q.get("query", "") for q in raw.get("related_searches", []) if q.get("query")]

    ai_overview = raw.get("ai_overview")
    ai_present = bool(ai_overview)
    ai_summary = ""
    if ai_overview:
        blocks = ai_overview.get("text_blocks", [])
        parts = []
        for b in blocks:
            if b.get("type") == "paragraph" and b.get("snippet"):
                parts.append(b["snippet"])
            elif b.get("type") == "list":
                for li in b.get("list", []):
                    if li.get("snippet"):
                        parts.append("- " + li["snippet"])
        ai_summary = "\n".join(parts) if parts else "(AI Overview présent mais texte non extrait par SerpAPI)"

    kg_present = bool(raw.get("knowledge_graph"))

    return {
        "keyword": keyword,
        "market": "FR (google.fr)",
        "organic": organic,
        "paa": paa,
        "related_searches": related_searches,
        "ai_overview_present": ai_present,
        "ai_overview_summary": ai_summary,
        "knowledge_graph_present": kg_present,
    }


def main():
    parser = argparse.ArgumentParser(description="Fetch live Google SERP (FR) via SerpAPI")
    parser.add_argument("--keyword", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--api-key", help="Overrides SERPAPI_API_KEY env var")
    args = parser.parse_args()

    api_key = args.api_key or os.environ.get("SERPAPI_API_KEY")
    if not api_key:
        sys.exit("Missing SerpAPI key: set SERPAPI_API_KEY in .env or pass --api-key")

    raw = fetch_serp(args.keyword, api_key)
    data = build_serp_json(raw, args.keyword)

    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    print(f"Saved: {args.out} ({len(data['organic'])} résultats organiques, "
          f"{len(data['paa'])} PAA, {len(data['related_searches'])} recherches associées)")


if __name__ == "__main__":
    main()
