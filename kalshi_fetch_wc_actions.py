import json,csv,time,requests
from datetime import datetime,timezone

BASE_URL="https://api.kalshi.com/trade-api/v2"
WC_SERIES=["KXMENWORLDCUP","KXWCGOALLEADER","KXWCGROUPWIN","KXWCFIRSTGOAL","KXWCTEAMFIRSTGOAL","KXWCSCORE","KXWC1HSCORE","KXWCGOAL","KXWCSOA","KXWCAST","KXWCCORNERS","KXWCTCORNERS","KXWC2HTOTAL","KXWC2H","KXWC2HSPREAD","KXWC2HBTTS","KXWCFTTS","KXWCTEAMTOTAL","KXWCBTTS","KXWCTOTAL","KXWCSPREAD","KXWC1H","KXWC1HTOTAL","KXWC1HSPREAD","KXWC1HBTTS","KXWCGAME","KXWCADVANCE","KXWCMATCH"]
CSV_FIELDS=["ticker","event_ticker","series_ticker","title","subtitle","category","status","close_time","yes_bid","yes_ask","no_bid","no_ask","last_price","volume","volume_24h","open_interest","liquidity","notional_value","yes_mid","prob_implied","spread","fetched_at"]
HEADERS={"Accept":"application/json"}

def fetch_series(s):
    mkts,cursor,page=[],None,0
    while True:
        params={"series_ticker":s,"limit":200}
        if cursor:params["cursor"]=cursor
        try:
            r=requests.get(f"{BASE_URL}/markets",params=params,headers=HEADERS,timeout=20)
            r.raise_for_status()
        except Exception as e:
            print(f"  [ERR] {s} p{page}: {e}");break
        data=r.json()
        batch=data.get("markets",[])
        mkts+=batch
        cursor=data.get("cursor")
        page+=1
        print(f"  {s} p{page}: +{len(batch)} (total {len(mkts)})")
        if not cursor or not batch:break
        time.sleep(0.2)
    return mkts

def enrich(m):
    m=dict(m)
    yb,ya=m.get("yes_bid"),m.get("yes_ask")
    if yb is not None and ya is not None:
        m["yes_mid"]=round((yb+ya)/2,2)
        m["prob_implied"]=round(m["yes_mid"]/100,4)
        m["spread"]=round(ya-yb,2)
    else:
        m["yes_mid"]=m["prob_implied"]=m["spread"]=None
    m["fetched_at"]=datetime.now(timezone.utc).isoformat()
    return m

def main():
    print(f"Iniciando {datetime.now(timezone.utc).isoformat()}")
    seen,all_mkts=set(),[]
    for s in WC_SERIES:
        print(f"-> {s}")
        for m in fetch_series(s):
            tk=m.get("ticker","")
            if tk and tk not in seen:
                seen.add(tk)
                all_mkts.append(enrich(m))
    print(f"\nTotal: {len(all_mkts)} mercados")
    with open("kalshi_wc_markets.json","w",encoding="utf-8") as f:
        json.dump({"fetched_at":datetime.now(timezone.utc).isoformat(),"total":len(all_mkts),"markets":all_mkts},f,indent=2,ensure_ascii=False)
    print("ok JSON")
    with open("kalshi_wc_markets.csv","w",newline="",encoding="utf-8") as f:
        w=csv.DictWriter(f,fieldnames=CSV_FIELDS,extrasaction="ignore")
        w.writeheader();w.writerows(all_mkts)
    print("ok CSV")
    top=sorted([m for m in all_mkts if m.get("volume")],key=lambda x:x["volume"],reverse=True)[:10]
    for m in top:
        prob=f"{m['prob_implied']*100:.1f}%" if m.get("prob_implied") is not None else "-"
        print(f"{m['ticker']:<40} {prob:>6} vol={m.get('volume',0):,}")

if __name__=="__main__":
    main()
