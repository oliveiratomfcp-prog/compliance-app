"""
fetch_data.py — Récupère les données ComplyAdvantage et génère data.json
Usage : python fetch_data.py [--days 30]
"""

import requests
import json
import time
import argparse
from datetime import datetime, timedelta
from pathlib import Path

# ── CONFIG ──────────────────────────────────────────
CA_API_KEY  = "JvqJdfYkae0ocShFbtU7Zv4ZXJY7JuzJ"
CA_BASE     = "https://api.complyadvantage.com"
PAGE_SIZE   = 100
OUTPUT_FILE = Path(__file__).parent / "data.json"
# ────────────────────────────────────────────────────

HEADERS = {"Authorization": f"Token {CA_API_KEY}"}
VERIFY_SSL = False

import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)


def fetch_searches(days: int) -> list:
    since = (datetime.utcnow() - timedelta(days=days)).strftime("%Y-%m-%d")
    results = []
    offset  = 0
    total   = None

    print(f"  Récupération des screenings depuis {since}...")

    while True:
        url = (
            f"{CA_BASE}/searches"
            f"?limit={PAGE_SIZE}&offset={offset}"
            f"&created_at_from={since}&ordering=-created_at"
        )
        resp = requests.get(url, headers=HEADERS, timeout=30, verify=VERIFY_SSL)

        if resp.status_code == 429:
            wait = 4
            print(f"  Rate limit atteint, pause {wait}s...")
            time.sleep(wait)
            continue

        resp.raise_for_status()
        content = resp.json().get("content", {})

        if total is None:
            total = content.get("total_count", 0)
            print(f"  Total à récupérer : {total}")

        batch = content.get("data", [])
        results.extend(batch)

        pct = int(len(results) / total * 100) if total else 100
        print(f"  [{pct:3d}%] {len(results)}/{total} screenings chargés", end="\r")

        if len(batch) < PAGE_SIZE or len(results) >= total:
            break

        offset += PAGE_SIZE

        # Pause tous les 500 pour respecter le rate limit
        if offset % 500 == 0:
            time.sleep(0.5)

    print(f"\n  ✓ {len(results)} screenings récupérés")
    return results


def fetch_users() -> dict:
    print("  Récupération des utilisateurs...")
    try:
        resp = requests.get(f"{CA_BASE}/users", headers=HEADERS, timeout=15, verify=VERIFY_SSL)
        resp.raise_for_status()
        users_list = resp.json().get("content", {}).get("data", [])
        users = {u["id"]: u for u in users_list}
        print(f"  ✓ {len(users)} utilisateurs récupérés")
        return users
    except Exception as e:
        print(f"  ⚠ Impossible de récupérer les utilisateurs : {e}")
        return {}


def slim_search(s: dict) -> dict:
    """Garde uniquement les champs utiles pour alléger le JSON."""
    term = s.get("search_term", "")
    if isinstance(term, dict):
        parts = [term.get("first_name",""), term.get("last_name","")]
        term  = " ".join(p for p in parts if p)

    return {
        "id":                s.get("id"),
        "search_term":       term,
        "client_ref":        s.get("client_ref"),
        "created_at":        s.get("created_at"),
        "status":            s.get("status"),
        "match_status":      s.get("match_status"),
        "total_hits":        s.get("total_hits", 0),
        "is_monitored":      s.get("is_monitored", False),
        "assigned_user_id":  s.get("assigned_user_id"),
        "filters":           {"types": s.get("filters", {}).get("types", [])},
        "hits":              s.get("hits", {}),
    }


def main():
    parser = argparse.ArgumentParser(description="Génère data.json pour le dashboard ComplyAdvantage")
    parser.add_argument("--days", type=int, default=30, help="Nombre de jours à récupérer (défaut: 30)")
    args = parser.parse_args()

    print(f"\n{'='*50}")
    print(f"  ComplyAdvantage — Génération data.json")
    print(f"  Période : {args.days} derniers jours")
    print(f"{'='*50}\n")

    users   = fetch_users()
    searches = fetch_searches(args.days)

    slimmed = [slim_search(s) for s in searches]

    output = {
        "generated_at": datetime.utcnow().isoformat() + "Z",
        "days":         args.days,
        "total":        len(slimmed),
        "users":        users,
        "searches":     slimmed,
    }

    OUTPUT_FILE.write_text(json.dumps(output, ensure_ascii=False, indent=2), encoding="utf-8")

    size_kb = OUTPUT_FILE.stat().st_size / 1024
    print(f"\n{'='*50}")
    print(f"  ✓ data.json généré ({size_kb:.0f} Ko)")
    print(f"  Fichier : {OUTPUT_FILE}")
    print(f"{'='*50}\n")
    print("  → Prochaine étape : git add data.json && git commit -m 'data: refresh' && git push")
    print()


if __name__ == "__main__":
    main()
