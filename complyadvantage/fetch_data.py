"""
fetch_data.py - Recupere TOUS les screenings ComplyAdvantage et genere data.json
Usage : py fetch_data.py
"""

import requests
import json
import time
import random
import urllib3
from datetime import datetime, timezone
from pathlib import Path

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ── CONFIG ─────────────────────────────────────────
CA_API_KEY  = "JvqJdfYkae0ocShFbtU7Zv4ZXJY7JuzJ"
CA_BASE     = "https://api.complyadvantage.com"
PAGE_SIZE   = 100
OUTPUT_FILE = Path(__file__).parent / "data.json"
VERIFY_SSL  = False
# ───────────────────────────────────────────────────

HEADERS = {"Authorization": f"Token {CA_API_KEY}"}


def api_get(url):
    """Appel GET avec retry exponentiel sur 429."""
    backoff = 2
    for attempt in range(8):
        try:
            resp = requests.get(url, headers=HEADERS, timeout=30, verify=VERIFY_SSL)
            if resp.status_code == 429:
                wait = backoff + random.random()
                print(f"\n  Rate limit (429), attente {wait:.1f}s...")
                time.sleep(wait)
                backoff = min(backoff * 2, 60)
                continue
            resp.raise_for_status()
            return resp.json()
        except requests.exceptions.RequestException as e:
            if attempt < 7:
                wait = backoff + random.random()
                print(f"\n  Erreur reseau, retry dans {wait:.1f}s : {e}")
                time.sleep(wait)
                backoff = min(backoff * 2, 60)
            else:
                raise


def fetch_all_searches():
    results = []
    offset  = 0
    total   = None

    print("  Recuperation de tous les screenings (sans filtre de date)...")

    while True:
        url  = f"{CA_BASE}/searches?limit={PAGE_SIZE}&offset={offset}&ordering=-created_at"
        data = api_get(url)
        content = data.get("content", {})

        if total is None:
            total = content.get("total_count", 0)
            print(f"  Total a recuperer : {total:,}\n")

        batch = content.get("data", [])
        results.extend(batch)

        pct = int(len(results) / total * 100) if total else 100
        print(f"  [{pct:3d}%] {len(results):,} / {total:,} screenings", end="\r")

        if len(batch) < PAGE_SIZE or len(results) >= total:
            break

        offset += PAGE_SIZE

        # Petite pause tous les 500 appels pour le rate limit
        if offset % 5000 == 0:
            time.sleep(0.5)

    print(f"\n  OK {len(results):,} screenings recuperes")
    return results


def fetch_users():
    print("  Recuperation des utilisateurs...")
    try:
        data  = api_get(f"{CA_BASE}/users")
        users = {u["id"]: u for u in data.get("content", {}).get("data", [])}
        print(f"  OK {len(users)} utilisateurs")
        return users
    except Exception as e:
        print(f"  Impossible de recuperer les utilisateurs : {e}")
        return {}


def slim(s):
    """Garde uniquement les champs utiles pour le dashboard."""
    term = s.get("search_term", "")
    if isinstance(term, dict):
        parts = [term.get("first_name", ""), term.get("last_name", "")]
        term  = " ".join(p for p in parts if p).strip()

    # Tags : cle->valeur, on les normalise en liste de strings "cle:valeur"
    raw_tags = s.get("tags") or {}
    if isinstance(raw_tags, dict):
        tags_list = [f"{k}:{v}" for k, v in raw_tags.items()]
    elif isinstance(raw_tags, list):
        tags_list = raw_tags
    else:
        tags_list = []

    return {
        "id":                             s.get("id"),
        "search_term":                    term,
        "client_ref":                     s.get("client_ref") or "",
        "created_at":                     s.get("created_at"),
        "status":                         s.get("status"),
        "match_status":                   s.get("match_status"),
        "risk_level":                     s.get("risk_level"),
        "total_hits":                     s.get("total_hits", 0),
        "is_monitored":                   bool(s.get("is_monitored")),
        "has_ongoing_monitoring_results": bool(s.get("has_ongoing_monitoring_results")),
        "assigned_to":                    s.get("assigned_to") or "",
        "assigned_user_id":               s.get("assigned_user_id"),
        "tags":                           tags_list,
        "filters":                        {"types": (s.get("filters") or {}).get("types", [])},
    }


def main():
    print(f"\n{'='*52}")
    print(f"  ComplyAdvantage - Generation data.json")
    print(f"  Tous les screenings (sans filtre de date)")
    print(f"{'='*52}\n")

    users    = fetch_users()
    searches = fetch_all_searches()
    slimmed  = [slim(s) for s in searches]

    output = {
        "generated_at": datetime.now(timezone.utc).isoformat(),
        "total":        len(slimmed),
        "users":        users,
        "searches":     slimmed,
    }

    OUTPUT_FILE.write_text(
        json.dumps(output, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )

    size_kb = OUTPUT_FILE.stat().st_size / 1024
    print(f"\n{'='*52}")
    print(f"  OK data.json genere ({size_kb:.0f} Ko, {len(slimmed):,} screenings)")
    print(f"  Fichier : {OUTPUT_FILE}")
    print(f"{'='*52}\n")
    print("  Prochaine etape :")
    print("  git add complyadvantage/data.json")
    print("  git commit -m \"data: refresh ComplyAdvantage\"")
    print("  git push\n")


if __name__ == "__main__":
    main()
