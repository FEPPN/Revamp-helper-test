#!/usr/bin/env python3
"""Finds the exact target page for a keyword on papernest.com and on each
competitor domain, via SerpAPI `site:` searches. Fully automatic, no LLM —
it takes the top-ranking result per domain and lists any other papernest.com
URLs seen in the top 10 as a cannibalization flag for a human to eyeball.

Usage:
  python find_pages.py --keyword "avis edf" --out pages.json
"""

import argparse
import json
import os
import sys
from urllib.parse import urlparse

import requests

try:
    from dotenv import load_dotenv
    load_dotenv()
except ImportError:
    pass

SERPAPI_URL = "https://serpapi.com/search.json"

TARGET_DOMAIN = "papernest.com"
COMPETITOR_DOMAINS = [
    "selectra.info",
    "kelwatt.fr",
    "fournisseurs-electricite.com",
    "hellowatt.fr",
]


def search_site_free(keyword, domain):
    """Free site-search fallback via DuckDuckGo's HTML endpoint (no API key,
    no quota) - only tried if SerpAPI errors out (quota exhausted, key
    invalid, etc.). Same technique validated in page-audit-pack. Returns a
    list shaped like SerpAPI's organic_results (just "link"/"title", the
    only fields this pipeline actually reads) - empty list if nothing
    usable comes back. Best-effort: DuckDuckGo's own ranking on a `site:`
    query can occasionally differ from Google's, and its bot detection is
    inconsistent, so this is a fallback for availability, not a guaranteed
    equivalent to SerpAPI."""
    try:
        from bs4 import BeautifulSoup
    except ImportError:
        return []
    import time
    import urllib.parse
    import urllib.request

    query = f"{keyword} site:{domain}"
    url = "https://html.duckduckgo.com/html/?" + urllib.parse.urlencode({"q": query})
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                      "(KHTML, like Gecko) Chrome/128.0 Safari/537.36",
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "fr-FR,fr;q=0.9,en;q=0.8",
        "Referer": "https://duckduckgo.com/",
    }
    html = ""
    for _ in range(3):
        req = urllib.request.Request(url, headers=headers)
        try:
            with urllib.request.urlopen(req, timeout=20) as resp:
                html = resp.read().decode("utf-8", errors="replace")
        except Exception:
            html = ""
        if html and "anomaly" not in html.lower():
            break
        time.sleep(1.5)
    if not html:
        return []

    soup = BeautifulSoup(html, "html.parser")
    results = []
    for link in soup.find_all("a", class_="result__a"):
        href = link.get("href", "")
        if "uddg=" in href:
            parsed = urllib.parse.urlparse(href)
            qs = urllib.parse.parse_qs(parsed.query)
            href = qs.get("uddg", [href])[0]
        if domain not in href:
            continue
        results.append({"link": href, "title": link.get_text(strip=True), "position": len(results) + 1})
        if len(results) >= 10:
            break
    return results


def search_site(keyword, domain, api_key):
    """Real Google via SerpAPI first (reliable); falls back to free
    DuckDuckGo only if SerpAPI errors out or its quota is exhausted, instead
    of crashing the whole app - matches the SerpAPI-first / free-fallback
    pattern used in page-audit-pack, for the same reason (a free-tier quota
    running out shouldn't take the whole tool down)."""
    params = {
        "engine": "google",
        "q": f"site:{domain} {keyword}",
        "gl": "fr",
        "hl": "fr",
        "num": 10,
        "api_key": api_key,
    }
    try:
        r = requests.get(SERPAPI_URL, params=params, timeout=30)
        r.raise_for_status()
        return r.json().get("organic_results", [])
    except requests.exceptions.RequestException:
        return search_site_free(keyword, domain)


def domain_of(url):
    return urlparse(url).netloc.replace("www.", "")


def main():
    parser = argparse.ArgumentParser(description="Find target + competitor pages for a keyword")
    parser.add_argument("--keyword", required=True)
    parser.add_argument("--out", required=True)
    parser.add_argument("--api-key", help="Overrides SERPAPI_API_KEY env var")
    args = parser.parse_args()

    api_key = args.api_key or os.environ.get("SERPAPI_API_KEY")
    if not api_key:
        sys.exit("Missing SerpAPI key: set SERPAPI_API_KEY in .env or pass --api-key")

    result = {"keyword": args.keyword, "target": None, "target_alternatives": [], "competitors": []}

    target_results = search_site(args.keyword, TARGET_DOMAIN, api_key)
    if target_results:
        result["target"] = {
            "url": target_results[0]["link"],
            "title": target_results[0].get("title", ""),
        }
        # Anything else from papernest.com in the top 10 = possible cannibalization
        for item in target_results[1:]:
            if TARGET_DOMAIN in domain_of(item["link"]):
                result["target_alternatives"].append({
                    "url": item["link"],
                    "title": item.get("title", ""),
                    "position": item.get("position"),
                })
    else:
        result["target"] = None

    for domain in COMPETITOR_DOMAINS:
        comp_results = search_site(args.keyword, domain, api_key)
        if comp_results:
            result["competitors"].append({
                "site": domain,
                "url": comp_results[0]["link"],
                "title": comp_results[0].get("title", ""),
                "no_dedicated_avis_page": False,
            })
        else:
            result["competitors"].append({
                "site": domain,
                "url": None,
                "title": None,
                "no_dedicated_avis_page": True,
            })

    with open(args.out, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"Saved: {args.out}")
    print(f"Cible : {result['target']['url'] if result['target'] else 'AUCUNE — vérifier le mot-clé'}")
    if result["target_alternatives"]:
        print(f"⚠️  Cannibalisation possible — autres pages papernest.com dans le top 10 : "
              f"{[a['url'] for a in result['target_alternatives']]}")
    for c in result["competitors"]:
        print(f"{c['site']}: {c['url'] or 'aucune page trouvée'}")


if __name__ == "__main__":
    main()
