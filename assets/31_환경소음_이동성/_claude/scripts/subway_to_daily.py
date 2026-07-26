# -*- coding: utf-8 -*-
# 지하철 승하차 → 서울 전체 일별 총승객(이동성 dose, 시간축). 25/26컬럼 2스키마 하모나이즈.
import csv, io, os
from collections import defaultdict

mob = r"C:\Users\wh850\AppData\Local\Temp\mobility"
files = {"2020":"subway_OA-12921_seq37_2020.csv","2021":"subway_seq38_2021.csv",
         "2022":"subway_seq39_2022.csv","2023":"subway_seq45_2023.csv"}

def read_text(p):
    for e in ("cp949","euc-kr","utf-8-sig","utf-8"):
        try:
            with open(p, encoding=e, newline="") as f: return f.read(), e
        except: pass
    with open(p, encoding="utf-8", errors="replace") as f: return f.read(), "utf-8-replace"

def norm_date(s):
    s=s.strip().replace(".","-").replace("/","-")[:10]
    dd=s.replace("-","")
    if len(dd)==8 and dd.isdigit(): return f"{dd[:4]}-{dd[4:6]}-{dd[6:8]}"
    return s

daily={}
for y,fn in files.items():
    txt,enc=read_text(os.path.join(mob,fn))
    rows=list(csv.reader(io.StringIO(txt))); h=rows[0]
    dci=next((i for i,c in enumerate(h) if ('날짜' in c) or ('수송일자' in c)), -1)
    META=('연번','날짜','수송일자','호선','역번호','고유역','외부역','역명','구분','역코드','등록','사용년월','노선','정류장','ID','ARS','합')
    tci=[i for i,c in enumerate(h) if i!=dci and not any(m in (c or '') for m in META)]
    print(f"   {y} timeCols 헤더 샘플:", [h[i] for i in tci[:2]]+['...']+[h[i] for i in tci[-2:]] if len(tci)>=4 else [h[i] for i in tci])
    agg=defaultdict(float); rc=defaultdict(int)
    for r in rows[1:]:
        if dci<0 or dci>=len(r): continue
        d=norm_date(r[dci]); s=0.0
        for i in tci:
            if i<len(r):
                v=r[i].strip().replace(",","")
                if v:
                    try: s+=float(v)
                    except: pass
        agg[d]+=s; rc[d]+=1
    print(f"{y}: enc={enc} cols={len(h)} dateCol={h[dci]!r} timeCols={len(tci)} days={len(agg)}")
    for d,v in agg.items(): daily[d]=v

out=os.path.join(mob,"subway_daily_seoul.csv")
with open(out,"w",newline="",encoding="utf-8-sig") as f:
    w=csv.writer(f); w.writerow(["date","subway_total"])
    for d in sorted(daily): w.writerow([d, int(daily[d])])
ks=sorted(daily)
print(f"\ndays={len(daily)} range {ks[0]}~{ks[-1]} -> subway_daily_seoul.csv")
print("sample:", [(k,int(daily[k])) for k in ks[:1]+ks[len(ks)//2:len(ks)//2+1]+ks[-1:]])
# 코로나 효과 sanity: 2019? 없음. 2020-02 vs 2020-04 비교(거리두기 충격)
for probe in ("2020-04-15","2021-07-15","2022-05-15","2023-05-15"):
    if probe in daily: print(f"  {probe}: {int(daily[probe]):,}")
print("=== DONE ===")
