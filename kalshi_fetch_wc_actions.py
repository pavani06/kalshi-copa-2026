"""
kalshi_fetch_wc_actions.py
Versão limpa para rodar no GitHub Actions (sem monkey-patching de DNS).
Grava kalshi_wc_markets.json e kalshi_wc_markets.csv na raiz do repo.
"""

import json
import csv
import time
import requests
from datetime import datetime, timezone

BASE_URL  = "https://api.kalshi.com/trade-api/v2"
WC_SERIES = [
    "KXMENWORLDCUP",
    "KXWCGOALLEADER",
    "KXWCMATCH",
    "KXWCGROUP",
    "KXWCADVANCE",
    "KXWCTOTAL",
]
CSV_FIELDS = [
    "ticker", "event_ticker", "series_ticker", "title",
    "subtitle", "category", "status", "close_time",
    "yes_bid", "yes_ask", "no_bid", "no_ask",
    "last_price", "volume", "volume_24h", "open_interest",
    "liquidity", "notional_value",
    "yes_mid", "prob_implied", "spread", "fetched_at",
]
HEADERS = {"Accept": "application/json"}


def fetch_markets_for_series(series_ticker: str) -> list[dict]:
    markets, cursor, page = [], None, 0
    while True:
        params = {"series_ticker": series_ticker, "limit": 200, "status": "open"}
        if cursor:
            params["cursor"] = cursor
        try:
            r = requests.get(f"{BASE_URL}/markets", params=params, headers=HEADERS, timeout=20)
            r.raise_for_status()
        except requests.RequestException as e:
            print(f"  [ERRO] {series_ticker} p{page}: {e}")
            break
        data     = r.json()
        batch    = data.get("markets", [])
        markets += batch
        cursor   = data.get("cursor")
        page    += 1
        print(f"  {series_ticker} p{page}: +{len(batch)} (total {len(markets)})")
        if not cursor or not batch:
            break
        time.sleep(0.25)
    return markets


def enrich(m: dict) -> dict:
    m = dict(m)
    yb, ya = m.get("yes_bid"), m.get("yes_ask")
    if yb is not None and ya is not None:
        m["yes_mid"]      = round((yb + ya) / 2, 2)
        m["prob_implied"] = round(m["yes_mid"] / 100, 4)
        m["spread"]       = round(ya - yb, 2)
    else:
        m["yes_mid"] = m["prob_implied"] = m["spread"] = None
    m["fetched_at"] = datetime.now(timezone.utc).isoformat()
    return m


def main():
    print(f"Iniciando {datetime.now(timezone.utc).isoformat()}")
    seen, all_mkts = set(), []
    for series in WC_SERIES:
        print(f"\n→ {series}")
        for m in fetch_markets_for_series(series):
            tk = m.get("ticker", "")
            if tk and tk not in seen:
                seen.add(tk)
                all_mkts.append(enrich(m))

    print(f"\nTotal: {len(all_mkts)} mercados")

    # JSON
    with open("kalshi_wc_markets.json", "w", encoding="utf-8") as f:
        json.dump({"fetched_at": datetime.now(timezone.utc).isoformat(),
                   "total": len(all_mkts), "markets": all_mkts},
                  f, indent=2, ensure_ascii=False)
    print("✓ kalshi_wc_markets.json")

    # CSV
    with open("kalshi_wc_markets.csv", "w", newline="", encoding="utf-8") as f:
        w = csv.DictWriter(f, fieldnames=CSV_FIELDS, extrasaction="ignore")
        w.writeheader()
        w.writerows(all_mkts)
    print("✓ kalshi_wc_markets.csv")

    # Sumário top 15 por volume
    top = sorted([m for m in all_mkts if m.get("volume")],
                 key=lambda x: x["volume"], reverse=True)[:15]
    print(f"\n{'TICKER':<28} {'PROB':>6} {'VOLUME':>10}")
    print("-" * 48)
    for m in top:
        prob = f"{m['prob_implied']*100:.1f}%" if m.get("prob_implied") is not None else "—"
        print(f"{m['ticker']:<28} {prob:>6} {m.get('volume',0):>10,}")


if __name__ == "__main__":
    main()
