#!/usr/bin/env python3
"""
gerar_dashboard.py — Regenera kalshi_dashboard.html com dados frescos
"""

import json, re, csv, os
from datetime import datetime, timezone

def load_json(path, default):
    try:
        with open(path, encoding='utf-8') as f:
            return json.load(f)
    except Exception as e:
        print(f"[WARN] {path}: {e}")
        return default

def load_kalshi():
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
            mid = round((bid+ask)/2 if bid and ask else lp, 2)
            if mid > 0:
                result[code] = {'bid': round(bid,2), 'ask': round(ask,2), 'mid': mid}
        except:
            pass
    print(f"  Kalshi: {len(result)} times — {list(result.keys())[:5]}")
    return result

def load_historico():
    raw = {}
    try:
        with open('historico_precos.csv', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f)
            by_day = {}
            for row in reader:
                date = row['data_hora'][:10]
                ticker = row['ticker']
                try:
                    val = float(row['kalshi_justa_pct'])
                    by_day[(date,ticker)] = val
                except:
                    pass
            # Agrupa por ticker em ordem de data
            from collections import defaultdict
            grouped = defaultdict(list)
            for (date,ticker),val in sorted(by_day.items()):
                grouped[ticker].append(val)
            raw = dict(grouped)
    except Exception as e:
        print(f"[WARN] historico_precos.csv: {e}")
    return raw

def load_news():
    data = load_json('news_events.json', [])
    filtered = [n for n in data if n.get('impact') in ('HIGH','MEDIUM') or n.get('source_type')=='price_move']
    return filtered[:40]

def load_book():
    BOOK = [
        {"sel":"BRASIL",    "tk":"BR","side":"S","qty":40, "avg":12.625,"flag":"🇧🇷"},
        {"sel":"ALEMANHA",  "tk":"DE","side":"L","qty":30, "avg":9.00,  "flag":"🇩🇪"},
        {"sel":"INGLATERRA","tk":"GB","side":"L","qty":20, "avg":11.75, "flag":"🏴󠁧󠁢󠁥󠁮󠁧󠁿"},
        {"sel":"PORTUGAL",  "tk":"PT","side":"L","qty":10, "avg":12.00, "flag":"🇵🇹"},
        {"sel":"MARROCOS",  "tk":"MA","side":"L","qty":30, "avg":2.42,  "flag":"🇲🇦"},
        {"sel":"COLOMBIA",  "tk":"CO","side":"L","qty":30, "avg":1.83,  "flag":"🇨🇴"},
        {"sel":"NORUEGA",   "tk":"NO","side":"L","qty":10, "avg":3.00,  "flag":"🇳🇴"},
    ]
    try:
        content = open('opcoes_model.py', encoding='utf-8').read()
        for p in BOOK:
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

def inject(html, kalshi, book, news, historico, ts):
    """Substitui blocos de dados — funciona com ou sem espaços ao redor do ="""

    # KALSHI — regex flexível
    html = re.sub(
        r'const KALSHI\s*=\s*\{[^;]+\};',
        f'const KALSHI={json.dumps(kalshi, ensure_ascii=False)};',
        html, flags=re.DOTALL
    )

    # BOOK
    html = re.sub(
        r'const BOOK\s*=\s*\[[^\]]+\];',
        f'const BOOK={json.dumps(book, ensure_ascii=False)};',
        html, flags=re.DOTALL
    )

    # RAW historico — pode ser grande, usa marcador
    raw_js = {}
    for tk, vals in historico.items():
        raw_js[tk] = [round(v,2) for v in vals[-8:]]
    html = re.sub(
        r'const RAW\s*=\s*\{[^;]+\};',
        f'const RAW={json.dumps(raw_js, ensure_ascii=False)};',
        html, flags=re.DOTALL
    )

    # NEWS_EMBEDDED — pode ter espaços
    html = re.sub(
        r'const NEWS_EMBEDDED\s*=\s*\[.*?\];',
        f'const NEWS_EMBEDDED={json.dumps(news, ensure_ascii=False)};',
        html, flags=re.DOTALL
    )

    # FETCH_TS — pode não existir, adiciona ou substitui
    if 'const FETCH_TS' in html:
        html = re.sub(
            r'const FETCH_TS\s*=\s*"[^"]*";',
            f'const FETCH_TS="{ts}";',
            html
        )
    else:
        # Injeta após KALSHI
        html = html.replace(
            f'const KALSHI={json.dumps(kalshi, ensure_ascii=False)};',
            f'const FETCH_TS="{ts}";\nconst KALSHI={json.dumps(kalshi, ensure_ascii=False)};'
        )

    return html

def main():
    ts = datetime.now(timezone.utc).strftime('%Y-%m-%dT%H:%M:%SZ')
    print(f"[gerar_dashboard] {ts}")

    kalshi    = load_kalshi()
    book      = load_book()
    news      = load_news()
    historico = load_historico()

    print(f"  Book:   {len(book)} posições")
    print(f"  News:   {len(news)} itens")
    print(f"  Hist:   {len(historico)} tickers")

    if not kalshi:
        print("[WARN] Kalshi vazio — abortando para não corromper HTML")
        return

    with open('kalshi_dashboard.html', encoding='utf-8') as f:
        html = f.read()

    html = inject(html, kalshi, book, news, historico, ts)

    # Verifica que substituições funcionaram
    checks = {
        'KALSHI': json.dumps(list(kalshi.keys())[:2])[1:-1].replace('"','') in html,
        'NEWS':   str(len(news)) in html,
        'RAW':    'const RAW=' in html,
    }
    for k,v in checks.items():
        print(f"  {'✅' if v else '❌'} {k}")

    with open('kalshi_dashboard.html', 'w', encoding='utf-8') as f:
        f.write(html)
    print(f"[OK] kalshi_dashboard.html — {len(html):,} chars")

main()
