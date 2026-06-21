#!/usr/bin/env python3
"""
gerar_dashboard.py — Regenera kalshi_dashboard.html com dados frescos
Roda via GitHub Actions a cada hora, após fetch_news.py e kalshi_fetch_wc_actions.py
"""

import json, re, csv, os
from datetime import datetime, timezone

# ── Lê dados frescos ──────────────────────────────────────────────────────────
def load_json(path, default):
    try:
        with open(path, encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"[WARN] {path}: {e}")
        return default

def load_kalshi():
    """Extrai mid prices dos mercados KXMENWORLDCUP-26-XX"""
    markets = load_json('kalshi_wc_markets.json', [])
    result = {}
    for m in markets:
        ticker = m.get('ticker','')
        if 'KXMENWORLDCUP-26-' not in ticker:
            continue
        code = ticker.replace('KXMENWORLDCUP-26-','')
        if len(code) > 4:
            continue
        try:
            bid = float(m.get('yes_bid_dollars') or 0) * 100
            ask = float(m.get('yes_ask_dollars') or 0) * 100
            lp  = float(m.get('last_price_dollars') or 0) * 100
            mid = (bid+ask)/2 if bid and ask else lp
            if mid > 0:
                result[code] = {'bid': round(bid,2), 'ask': round(ask,2), 'mid': round(mid,2)}
        except:
            pass
    return result

def load_historico():
    """Lê historico_precos.csv e monta série por ticker"""
    raw = {}
    try:
        with open('historico_precos.csv', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            by_day = {}
            for row in reader:
                date = row['data_hora'][:10]
                ticker = row['ticker']
                key = (date, ticker)
                try:
                    val = float(row['kalshi_justa_pct'])
                    by_day[key] = {'date': date, 'ticker': ticker, 'val': val, 'sel': row.get('selecao','')}
                except:
                    pass
            for (date, ticker), v in sorted(by_day.items()):
                if ticker not in raw:
                    raw[ticker] = []
                raw[ticker].append({'date': date, 'val': v['val']})
    except Exception as e:
        print(f"[WARN] historico_precos.csv: {e}")
    return raw

def load_news():
    """Carrega news_events.json, filtra HIGH/MEDIUM, máx 40"""
    data = load_json('news_events.json', [])
    filtered = [n for n in data if n.get('impact') in ('HIGH','MEDIUM') or n.get('source_type')=='price_move']
    return filtered[:40]

def load_book():
    """Lê book do opcoes_model.py"""
    BOOK = [
        {"sel":"BRASIL",    "tk":"BR","side":"S","qty":40, "avg":12.625,"flag":"🇧🇷"},
        {"sel":"ALEMANHA",  "tk":"DE","side":"L","qty":30, "avg":9.00,  "flag":"🇩🇪"},
        {"sel":"INGLATERRA","tk":"GB","side":"L","qty":20, "avg":11.75, "flag":"🏴󠁧󠁢󠁥󠁮󠁧󠁿"},
        {"sel":"PORTUGAL",  "tk":"PT","side":"L","qty":10, "avg":12.00, "flag":"🇵🇹"},
        {"sel":"MARROCOS",  "tk":"MA","side":"L","qty":30, "avg":2.42,  "flag":"🇲🇦"},
        {"sel":"COLOMBIA",  "tk":"CO","side":"L","qty":30, "avg":1.83,  "flag":"🇨🇴"},
        {"sel":"NORUEGA",   "tk":"NO","side":"L","qty":10, "avg":3.00,  "flag":"🇳🇴"},
        {"sel":"SENEGAL",   "tk":"SN","side":"L","qty":10, "avg":0.75,  "flag":"🇸🇳"},
    ]
    # Tenta ler posições atuais do opcoes_model.py
    try:
        content = open('opcoes_model.py', encoding='utf-8').read()
        for p in BOOK:
            # Procura linha com o time e extrai contratos e preço médio
            pattern = rf"'selecao':\s*'{p['sel'].capitalize()}[^']*'[^}}]*'contratos':\s*(\d+)[^}}]*'preco_medio':\s*([\d.]+)"
            # Tenta variante case-insensitive
            lines = [l for l in content.split('\n') if p['sel'] in l.upper() and 'contratos' in l]
            for line in lines:
                m_qty = re.search(r"'contratos':\s*(\d+)", line)
                m_avg = re.search(r"'preco_medio':\s*([\d.]+)", line)
                if m_qty and m_avg:
                    p['qty'] = int(m_qty.group(1))
                    p['avg'] = float(m_avg.group(1))
                    break
    except Exception as e:
        print(f"[WARN] opcoes_model.py: {e}")
    return BOOK

# ── Lê template HTML ──────────────────────────────────────────────────────────
def load_template():
    """Lê o HTML atual como template, substituindo apenas os blocos de dados"""
    with open('kalshi_dashboard.html', encoding='utf-8') as f:
        return f.read()

# ── Injeta dados no HTML ──────────────────────────────────────────────────────
def inject(html, kalshi, book, news, historico, ts):
    """Substitui blocos de dados dinâmicos no HTML"""

    # KALSHI data
    html = re.sub(
        r'const KALSHI = \{[^;]+\};',
        f'const KALSHI = {json.dumps(kalshi, ensure_ascii=False)};',
        html, flags=re.DOTALL
    )

    # BOOK data
    html = re.sub(
        r'const BOOK = \[[^\]]+\];',
        f'const BOOK = {json.dumps(book, ensure_ascii=False)};',
        html, flags=re.DOTALL
    )

    # NEWS_EMBEDDED
    html = re.sub(
        r'const NEWS_EMBEDDED = \[[^\]]*\];',
        f'const NEWS_EMBEDDED = {json.dumps(news, ensure_ascii=False)};',
        html, flags=re.DOTALL
    )

    # RAW histórico (apenas tickers do book)
    book_tickers = [p['tk'] for p in book]
    raw_js = {}
    for tk in book_tickers:
        if tk in historico:
            raw_js[tk] = [p['val'] for p in historico[tk][-8:]]  # últimos 8 dias
    # Adiciona tickers extras que existem no histórico
    for tk, pts in historico.items():
        if tk not in raw_js:
            raw_js[tk] = [p['val'] for p in pts[-8:]]
    html = re.sub(
        r'const RAW = \{[^;]+\};',
        f'const RAW = {json.dumps(raw_js, ensure_ascii=False)};',
        html, flags=re.DOTALL
    )

    # FETCH_TS
    html = re.sub(
        r'const FETCH_TS = "[^"]+";',
        f'const FETCH_TS = "{ts}";',
        html
    )

    return html

# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    ts = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
    print(f"[gerar_dashboard] {ts}")

    kalshi    = load_kalshi()
    book      = load_book()
    news      = load_news()
    historico = load_historico()

    print(f"  Kalshi: {len(kalshi)} times")
    print(f"  Book:   {len(book)} posições")
    print(f"  News:   {len(news)} itens")
    print(f"  Hist:   {len(historico)} tickers")

    if not kalshi:
        print("[WARN] Kalshi vazio — mantendo HTML atual")
        return

    html = load_template()
    html = inject(html, kalshi, book, news, historico, ts)

    with open('kalshi_dashboard.html', 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"[OK] kalshi_dashboard.html regenerado — {len(html):,} chars")

main()
