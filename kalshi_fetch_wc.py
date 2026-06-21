"""
kalshi_fetch_wc.py
Puxa TODOS os mercados da Copa do Mundo 2026 na Kalshi (sem autenticação),
com paginação completa, e grava:
  - kalshi_wc_markets.json   → dados brutos completos
  - kalshi_wc_markets.csv    → tabela plana para análise / Excel

Uso:
    pip install requests dnspython
    python kalshi_fetch_wc.py

Dependências: requests, dnspython
"""

import json
import csv
import time
import socket
import requests
from datetime import datetime, timezone

# ── DNS LEAK FIX ──────────────────────────────────────────────────────────────
# No Windows, o DNS muitas vezes não roteia pelo túnel VPN (DNS leak).
# Forçamos a resolução via Google (8.8.8.8) para bypassar o resolver local.
try:
    import dns.resolver as _dns_resolver

    _resolver = _dns_resolver.Resolver()
    _resolver.nameservers = ["8.8.8.8", "8.8.4.4"]
    _resolver.timeout = 5
    _resolver.lifetime = 10

    _orig_getaddrinfo = socket.getaddrinfo

    def _patched_getaddrinfo(host, port, family=0, type=0, proto=0, flags=0):
        try:
            answers = _resolver.resolve(host, "A")
            ip = answers[0].to_text()
            return _orig_getaddrinfo(ip, port, family, type, proto, flags)
        except Exception:
            return _orig_getaddrinfo(host, port, family, type, proto, flags)

    socket.getaddrinfo = _patched_getaddrinfo
    print("[DNS] Resolver forçado via 8.8.8.8 (fix DNS leak VPN)\n")

except ImportError:
    print("[AVISO] dnspython não instalado. Rodando com DNS padrão do sistema.")
    print("        Se falhar, rode: pip install dnspython\n")

# ── CONFIG ────────────────────────────────────────────────────────────────────
BASE_URL   = "https://api.kalshi.com/trade-api/v2"
# Tickers de séries da Copa 2026 conhecidas na Kalshi
WC_SERIES  = [
    "KXMENWORLDCUP",   # campeão geral
    "KXWCGOALLEADER",  # artilheiro / golden boot
    "KXWCMATCH",       # resultados jogo a jogo
    "KXWCGROUP",       # fase de grupos
    "KXWCADVANCE",     # quem avança
    "KXWCTOTAL",       # total de gols / props
]
OUTPUT_DIR = "."       # mesma pasta do script; ajuste se quiser subpasta

# Campos que queremos manter no CSV (subset útil para análise de prob)
CSV_FIELDS = [
    "ticker", "event_ticker", "series_ticker", "title",
    "subtitle", "category", "status", "close_time",
    "yes_bid", "yes_ask", "no_bid", "no_ask",
    "last_price", "volume", "volume_24h", "open_interest",
    "liquidity", "notional_value",
]

HEADERS = {"Accept": "application/json"}

# ── FETCH COM PAGINAÇÃO ───────────────────────────────────────────────────────
def fetch_markets_for_series(series_ticker: str) -> list[dict]:
    """Busca todos os mercados de uma série com paginação por cursor."""
    markets = []
    cursor  = None
    page    = 0

    while True:
        params: dict = {"series_ticker": series_ticker, "limit": 200, "status": "open"}
        if cursor:
            params["cursor"] = cursor

        try:
            r = requests.get(f"{BASE_URL}/markets", params=params, headers=HEADERS, timeout=15)
            r.raise_for_status()
        except requests.RequestException as e:
            print(f"  [ERRO] {series_ticker} página {page}: {e}")
            break

        data      = r.json()
        batch     = data.get("markets", [])
        markets  += batch
        cursor    = data.get("cursor")
        page     += 1

        print(f"  {series_ticker} p{page}: +{len(batch)} mercados (total {len(markets)})")

        if not cursor or not batch:
            break
        time.sleep(0.2)   # respeitar rate limit (5 req/s aprox.)

    return markets


def fetch_all_wc_markets() -> list[dict]:
    """Itera todas as séries e agrega sem duplicatas (por ticker)."""
    seen    = set()
    all_mkts = []

    for series in WC_SERIES:
        print(f"\n→ Buscando série: {series}")
        batch = fetch_markets_for_series(series)
        for m in batch:
            tk = m.get("ticker", "")
            if tk and tk not in seen:
                seen.add(tk)
                all_mkts.append(m)

    return all_mkts


# ── ENRICH: normalizar preços (Kalshi retorna em centavos inteiros 0-100) ─────
def enrich(market: dict) -> dict:
    """Adiciona campos derivados úteis."""
    m = dict(market)

    # Preço implícito = mid do yes_bid / yes_ask (probabilidade em %)
    yb = m.get("yes_bid")
    ya = m.get("yes_ask")
    if yb is not None and ya is not None:
        m["yes_mid"]      = round((yb + ya) / 2, 2)
        m["prob_implied"] = round(m["yes_mid"] / 100, 4)  # 0.0 – 1.0
    else:
        m["yes_mid"]      = None
        m["prob_implied"] = None

    # Spread
    if yb is not None and ya is not None:
        m["spread"] = round(ya - yb, 2)
    else:
        m["spread"] = None

    m["fetched_at"] = datetime.now(timezone.utc).isoformat()
    return m


# ── GRAVAR ────────────────────────────────────────────────────────────────────
def save_json(markets: list[dict], path: str):
    with open(path, "w", encoding="utf-8") as f:
        json.dump({"fetched_at": datetime.now(timezone.utc).isoformat(),
                   "total": len(markets),
                   "markets": markets}, f, indent=2, ensure_ascii=False)
    print(f"\n✓ JSON gravado → {path}")


def save_csv(markets: list[dict], path: str):
    # Detecta todas as colunas disponíveis
    all_keys = list(dict.fromkeys(
        CSV_FIELDS + ["yes_mid", "prob_implied", "spread", "fetched_at"]
    ))

    with open(path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=all_keys, extrasaction="ignore")
        writer.writeheader()
        for m in markets:
            writer.writerow(m)
    print(f"✓ CSV gravado  → {path}")


# ── SUMÁRIO ───────────────────────────────────────────────────────────────────
def print_summary(markets: list[dict]):
    print(f"\n{'='*60}")
    print(f"  TOTAL DE MERCADOS: {len(markets)}")
    print(f"{'='*60}")

    # Top 20 por volume
    top = sorted(
        [m for m in markets if m.get("volume") is not None],
        key=lambda x: x["volume"], reverse=True
    )[:20]

    print(f"\n{'TICKER':<30} {'TÍTULO':<35} {'PROB':>6} {'VOL':>10}")
    print("-" * 85)
    for m in top:
        prob = f"{m.get('prob_implied', 0)*100:.1f}%" if m.get("prob_implied") is not None else "  —  "
        vol  = f"{m.get('volume', 0):,}"
        title = (m.get("title") or "")[:34]
        print(f"{m['ticker']:<30} {title:<35} {prob:>6} {vol:>10}")


# ── MAIN ──────────────────────────────────────────────────────────────────────
def main():
    print("Kalshi World Cup 2026 — Market Fetcher")
    print(f"Iniciando em {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")

    markets = fetch_all_wc_markets()
    markets = [enrich(m) for m in markets]

    if not markets:
        print("\n[AVISO] Nenhum mercado encontrado. Verifique os tickers de série.")
        return

    json_path = f"{OUTPUT_DIR}/kalshi_wc_markets.json"
    csv_path  = f"{OUTPUT_DIR}/kalshi_wc_markets.csv"

    save_json(markets, json_path)
    save_csv(markets, csv_path)
    print_summary(markets)

    print(f"\nFinalizado. {len(markets)} mercados gravados.")


if __name__ == "__main__":
    main()
