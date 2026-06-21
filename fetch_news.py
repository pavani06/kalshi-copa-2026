#!/usr/bin/env python3
"""
fetch_news.py — Busca notícias de impacto para o book da Copa 2026
Roda via GitHub Actions a cada hora.
"""

import json, re, time, hashlib, xml.etree.ElementTree as ET
from datetime import datetime, timezone
from urllib.request import urlopen, Request
from urllib.parse import quote_plus
from urllib.error import HTTPError

TEAMS = {
    "Brasil":     {"en": "Brazil",   "ticker": "BR", "side": "S"},
    "Alemanha":   {"en": "Germany",  "ticker": "DE", "side": "L"},
    "Inglaterra": {"en": "England",  "ticker": "GB", "side": "L"},
    "Portugal":   {"en": "Portugal", "ticker": "PT", "side": "L"},
    "Marrocos":   {"en": "Morocco",  "ticker": "MA", "side": "L"},
    "Colombia":   {"en": "Colombia", "ticker": "CO", "side": "L"},
    "Noruega":    {"en": "Norway",   "ticker": "NO", "side": "L"},
    "Senegal":    {"en": "Senegal",  "ticker": "SN", "side": "L"},
}

# Queries focadas em impacto de preço — máximo sinal, mínimo ruído
QUERIES = [
    # Lesões e suspensões por time
    *[{"q": f'{m["en"]} injury suspended World Cup 2026', "cat": "injury",  "teams": [pt], "impact": "HIGH"}
      for pt, m in TEAMS.items()],
    # Escalações confirmadas
    *[{"q": f'{m["en"]} lineup confirmed World Cup 2026', "cat": "lineup",  "teams": [pt], "impact": "HIGH"}
      for pt, m in TEAMS.items()],
    # Resultados de hoje
    {"q": "World Cup 2026 match result today score",        "cat": "result",  "teams": [], "impact": "HIGH"},
    {"q": "Copa do Mundo 2026 resultado hoje gol",          "cat": "result",  "teams": [], "impact": "HIGH"},
    # Chaveamento e odds
    {"q": "World Cup 2026 knockout bracket odds next round","cat": "bracket", "teams": [], "impact": "MEDIUM"},
    # Notícia geral copa
    {"q": "World Cup 2026 news today",                      "cat": "general", "teams": [], "impact": "LOW"},
]

IMPACT_KW = {
    "HIGH": ["injur","suspend","ban","out ","doubtful","lesão","suspenso","fora","lineup","confirmed","xi","escalação",
             "result","score","won","lost","beat","defeat","eliminat","qualif","gol","placar","venceu","perdeu",
             "red card","cartão vermelho","fracture","muscle","hamstring","knee"],
    "MEDIUM": ["odds","probability","chance","bracket","knockout","chaveamento","favorit","doubt","dúvida","treino","training"],
}

ALL_NAMES = {}
for pt, meta in TEAMS.items():
    ALL_NAMES[pt.lower()] = pt
    ALL_NAMES[meta["en"].lower()] = pt
    # aliases comuns
ALL_NAMES["brazil"] = "Brasil"
ALL_NAMES["england"] = "Inglaterra"
ALL_NAMES["germany"] = "Alemanha"
ALL_NAMES["morocco"] = "Marrocos"
ALL_NAMES["norway"] = "Noruega"
ALL_NAMES["colombia"] = "Colombia"

def detect_teams(text):
    tl = text.lower()
    return list({canonical for name, canonical in ALL_NAMES.items() if name in tl})

def classify_impact(text):
    tl = text.lower()
    for kw in IMPACT_KW["HIGH"]:
        if kw in tl: return "HIGH"
    for kw in IMPACT_KW["MEDIUM"]:
        if kw in tl: return "MEDIUM"
    return "LOW"

def classify_cat(text, default):
    t = text.lower()
    if any(k in t for k in ["injur","lesão","suspend","suspenso","ban","out","cartão","fracture","muscle","hamstring","knee"]): return "injury"
    if any(k in t for k in ["lineup","starting","xi","escalação","titular","confirmed","formation"]): return "lineup"
    if any(k in t for k in ["won","lost","beat","defeat","result","gol","score","placar","venceu","perdeu","eliminat","qualif"]): return "result"
    if any(k in t for k in ["bracket","knockout","chaveamento","round of","oitavas","quartas","semifinal"]): return "bracket"
    return default

UA_LIST = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.4 Safari/605.1.15",
    "Mozilla/5.0 (X11; Linux x86_64; rv:126.0) Gecko/20100101 Firefox/126.0",
]
ua_idx = 0

def fetch_rss(query, max_items=5):
    global ua_idx
    # Tenta Google News RSS com 2 variantes de URL
    urls = [
        f"https://news.google.com/rss/search?q={quote_plus(query)}&hl=en-US&gl=US&ceid=US:en",
        f"https://news.google.com/rss/search?q={quote_plus(query)}&hl=pt-BR&gl=BR&ceid=BR:pt-419",
    ]
    for url in urls:
        try:
            ua = UA_LIST[ua_idx % len(UA_LIST)]
            ua_idx += 1
            req = Request(url, headers={
                "User-Agent": ua,
                "Accept": "application/rss+xml,application/xml,text/xml,*/*",
                "Accept-Language": "en-US,en;q=0.9",
                "Cache-Control": "no-cache",
            })
            with urlopen(req, timeout=20) as r:
                xml_data = r.read()
            root = ET.fromstring(xml_data)
            items = []
            for item in root.findall(".//item")[:max_items]:
                title = item.findtext("title", "").strip()
                # Remove " - Source Name" suffix do Google News
                title = re.sub(r'\s+-\s+[^-]+$', '', title).strip()
                link  = item.findtext("link", "").strip()
                pub   = item.findtext("pubDate", "").strip()
                source = item.findtext("source", "")
                try:
                    from email.utils import parsedate_to_datetime
                    dt = parsedate_to_datetime(pub)
                    ts = dt.timestamp()
                    date_str = dt.strftime("%d/%m %H:%M")
                except:
                    ts = time.time()
                    date_str = datetime.now().strftime("%d/%m %H:%M")
                if title:
                    items.append({"title": title, "url": link, "source": source,
                                  "date": date_str, "date_ts": ts})
            if items:
                return items
        except HTTPError as e:
            if e.code == 429:
                print(f"  [429] Rate limit — aguardando 5s")
                time.sleep(5)
            else:
                print(f"  [HTTP {e.code}] {query[:40]}")
        except Exception as e:
            print(f"  [ERR] {query[:40]}: {type(e).__name__}: {e}")
        time.sleep(1)
    return []

def main():
    print(f"[fetch_news] {datetime.now(timezone.utc).strftime('%Y-%m-%d %H:%M')} UTC")

    try:
        with open("news_events.json", encoding="utf-8") as f:
            existing = json.load(f)
    except:
        existing = []

    # Remove itens >7 dias
    cutoff = time.time() - 7 * 86400
    existing = [n for n in existing if n.get("date_ts", 0) > cutoff]

    existing_urls   = {n.get("url","") for n in existing if n.get("url")}
    existing_hashes = {hashlib.md5(n.get("title","").encode()).hexdigest() for n in existing}
    seen_run = set()
    new_items = []

    for q_meta in QUERIES:
        print(f"  [{q_meta['cat'].upper():8}] {q_meta['q'][:55]}")
        items = fetch_rss(q_meta["q"], max_items=4)
        for item in items:
            h = hashlib.md5(item["title"].encode()).hexdigest()
            if item["url"] in existing_urls or h in existing_hashes or h in seen_run:
                continue
            seen_run.add(h)
            teams  = detect_teams(item["title"]) or q_meta["teams"]
            impact = classify_impact(item["title"])
            cat    = classify_cat(item["title"], q_meta["cat"])
            # Filtra baixo impacto sem times do book
            if impact == "LOW" and not teams and cat not in ["result","bracket"]:
                continue
            new_items.append({
                "date": item["date"], "date_ts": item["date_ts"],
                "title": item["title"], "source": item["source"],
                "url": item["url"], "teams": teams,
                "category": cat, "impact": impact,
                "source_type": "news",
            })
        time.sleep(0.8)

    print(f"\n[OK] Novos: {len(new_items)} | Existentes: {len(existing)}")
    all_items = (new_items + existing)
    all_items.sort(key=lambda x: x.get("date_ts", 0), reverse=True)
    all_items = all_items[:60]

    with open("news_events.json", "w", encoding="utf-8") as f:
        json.dump(all_items, f, ensure_ascii=False, indent=2)
    print(f"[OK] Salvo: {len(all_items)} itens")

main()
