#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
salvar_snapshot.py — arquiva snapshots datados (Kalshi + Call da Copa)
=====================================================================
A cada execucao guarda o estado atual dos dois mercados:

  historico/kalshi_AAAA-MM-DD_HHMM.json   (prob justa por selecao)
  historico/call_AAAA-MM-DD_HHMM.json     (preco do Call da Copa)
  historico_precos.csv                    (1 linha por selecao por data;
                                           formato longo p/ grafico/pivot)

Colunas do CSV: data_hora, ticker, selecao, kalshi_justa_pct,
                call_preco_pct, edge_pp

Uso:
  python salvar_snapshot.py
Rode depois de atualizar (kalshi_update.py / ler_call_da_copa.py).
"""

import csv
import json
import os
import datetime as dt

HERE = os.path.dirname(os.path.abspath(__file__))
SNAP = os.path.join(HERE, "kalshi_snapshot.json")
CALL = os.path.join(HERE, "call_prices.json")
HIST = os.path.join(HERE, "historico")
CSV = os.path.join(HERE, "historico_precos.csv")


def load(path):
    try:
        return json.load(open(path, encoding="utf-8"))
    except Exception:
        return {}


def main():
    os.makedirs(HIST, exist_ok=True)
    now = dt.datetime.now()
    stamp_file = now.strftime("%Y-%m-%d_%H%M")
    stamp_iso = now.strftime("%Y-%m-%d %H:%M")

    snap = load(SNAP)
    call = load(CALL).get("prices", {})

    if not snap.get("markets"):
        print("[x] kalshi_snapshot.json vazio. Rode kalshi_update.py antes.")
        return

    # 1) copias datadas (fidelidade total)
    json.dump(snap, open(os.path.join(HIST, f"kalshi_{stamp_file}.json"), "w",
              encoding="utf-8"), ensure_ascii=False, indent=2)
    json.dump({"updated": stamp_iso, "prices": call},
              open(os.path.join(HIST, f"call_{stamp_file}.json"), "w",
              encoding="utf-8"), ensure_ascii=False, indent=2)

    # 2) linhas para o CSV historico (1 por selecao)
    new = not os.path.exists(CSV)
    n = 0
    with open(CSV, "a", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f)
        if new:
            w.writerow(["data_hora", "ticker", "selecao", "kalshi_justa_pct",
                        "call_preco_pct", "edge_pp"])
        for m in snap["markets"]:
            tk = m["ticker"]
            kj = round(m["prob_justa"] * 100, 2)
            cp = call.get(tk)
            edge = round(cp - kj, 2) if cp is not None else ""
            w.writerow([stamp_iso, tk.replace("KXMENWORLDCUP-26-", ""),
                        m.get("pt") or m.get("country"), kj,
                        cp if cp is not None else "", edge])
            n += 1

    ncall = sum(1 for m in snap["markets"] if m["ticker"] in call)
    print(f"[ok] Snapshot Kalshi -> historico/kalshi_{stamp_file}.json")
    print(f"[ok] Snapshot Call   -> historico/call_{stamp_file}.json ({len(call)} precos)")
    print(f"[ok] historico_precos.csv += {n} linhas ({ncall} com preco do Call).")


if __name__ == "__main__":
    main()
