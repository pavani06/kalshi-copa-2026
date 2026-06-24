#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
salvar_snapshot.py — anexa o snapshot diário ao historico_precos.csv
====================================================================
Lê `kalshi_wc_markets.json` (gerado pelo passo "Fetch mercados Kalshi"
do workflow) e grava 1 linha por seleção no histórico, que alimenta os
gráficos do dashboard (const RAW).

Correção (24/jun/2026): a versão anterior lia `kalshi_snapshot.json`,
arquivo que o pipeline do Actions nunca produz — então abortava em todo
run e o histórico congelou em 20/jun. Agora lê a mesma fonte que o
`gerar_dashboard.py` já consome.

Colunas do CSV: data_hora, ticker, selecao, kalshi_justa_pct,
                call_preco_pct, edge_pp
"""

import csv
import json
import os
import datetime as dt

HERE = os.path.dirname(os.path.abspath(__file__))
KWM  = os.path.join(HERE, "kalshi_wc_markets.json")   # fonte real (Actions)
CALL = os.path.join(HERE, "call_prices.json")          # opcional (Call da Copa)
HIST = os.path.join(HERE, "historico")
CSV  = os.path.join(HERE, "historico_precos.csv")
PREF = "KXMENWORLDCUP-26-"


def load_json(path, default):
    try:
        with open(path, encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def name_map_from_csv():
    """Reaproveita os nomes PT já presentes no histórico p/ manter a
    coluna `selecao` consistente (o arquivo Kalshi traz nomes em inglês)."""
    m = {}
    if os.path.exists(CSV):
        try:
            with open(CSV, encoding="utf-8-sig") as f:
                for row in csv.DictReader(f):
                    tk, sel = row.get("ticker"), row.get("selecao")
                    if tk and sel:
                        m[tk] = sel
        except Exception:
            pass
    return m


def mid_pct(m):
    """Probabilidade justa (%) = mid do yes; fallback no last price."""
    try:
        yb = float(m.get("yes_bid_dollars") or 0) * 100
        ya = float(m.get("yes_ask_dollars") or 0) * 100
        lp = float(m.get("last_price_dollars") or 0) * 100
        return round((yb + ya) / 2 if (yb and ya) else lp, 2)
    except Exception:
        return None


def main():
    os.makedirs(HIST, exist_ok=True)
    now = dt.datetime.now(dt.timezone.utc)
    stamp_file = now.strftime("%Y-%m-%d_%H%M")
    stamp_iso = now.strftime("%Y-%m-%d %H:%M")

    markets = load_json(KWM, [])
    if not markets:
        print("[x] kalshi_wc_markets.json vazio/ausente. Abortando sem tocar no CSV.")
        return

    call_raw = load_json(CALL, {})
    call = call_raw.get("prices", {}) if isinstance(call_raw, dict) else {}
    names = name_map_from_csv()

    rows, archive = [], {}
    for m in markets:
        tk = m.get("ticker", "")
        if PREF not in tk:
            continue
        code = tk.replace(PREF, "")
        if len(code) > 4:
            continue
        kj = mid_pct(m)
        if kj is None or kj <= 0:
            continue
        team = names.get(code) or m.get("yes_sub_title") or m.get("no_sub_title") or code
        cp = call.get(code)
        edge = round(cp - kj, 2) if cp is not None else ""
        rows.append([stamp_iso, code, team, kj, cp if cp is not None else "", edge])
        archive[code] = kj

    if not rows:
        print("[x] Nenhum market WC válido. Abortando.")
        return

    # cópia datada p/ auditoria
    json.dump(
        {"updated": stamp_iso, "kalshi_pct": archive},
        open(os.path.join(HIST, f"kalshi_{stamp_file}.json"), "w", encoding="utf-8"),
        ensure_ascii=False, indent=2,
    )

    # append no CSV (utf-8 puro: o arquivo já tem BOM no início; não reinjetar)
    new = not os.path.exists(CSV)
    with open(CSV, "a", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        if new:
            f.write("\ufeff")  # BOM só se o arquivo for novo
            w.writerow(["data_hora", "ticker", "selecao",
                        "kalshi_justa_pct", "call_preco_pct", "edge_pp"])
        w.writerows(rows)

    ncall = sum(1 for r in rows if r[4] != "")
    print(f"[ok] historico_precos.csv += {len(rows)} linhas @ {stamp_iso} "
          f"({ncall} com preço do Call)")


if __name__ == "__main__":
    main()
