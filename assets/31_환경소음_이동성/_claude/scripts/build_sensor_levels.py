# -*- coding: utf-8 -*-
# 서울 전역 S-DoT 센서별 LAeq 레벨(연/분기/주야) 산출 — 공간검증 + 본연구 재사용
import os, zipfile, io, csv, math
from collections import defaultdict
import openpyxl

tmp=r"C:\Users\wh850\AppData\Local\Temp"; sdot_dir=os.path.join(tmp,"sdot")
def decode(b):
    for e in ("cp949","euc-kr","utf-8-sig","utf-8"):
        try: return b.decode(e)
        except: pass
    return b.decode("utf-8","replace")
def cidx(h,p):
    for i,x in enumerate(h):
        if p in (x or ""): return i
    return -1

wb=openpyxl.load_workbook(os.path.join(sdot_dir,"location.xlsx"),read_only=True)
rows=list(wb.active.iter_rows(values_only=True)); hdr=[str(x) if x is not None else "" for x in rows[0]]
lat_i=cidx(hdr,"위도"); lon_i=cidx(hdr,"경도"); ser2c={}
for r in rows[1:]:
    if r[1]:
        try: ser2c[str(r[1]).strip()]=(float(r[lat_i]),float(r[lon_i]))
        except: pass
print("S-DoT coords:", len(ser2c))

def build(zname, outcsv):
    acc=defaultdict(lambda: defaultdict(lambda:[0.0,0]))  # serial -> key -> [sum_linear, n]
    z=zipfile.ZipFile(os.path.join(sdot_dir,zname))
    files=[x for x in z.namelist() if x.lower().endswith(".csv")]
    for n in files:
        rdr=csv.reader(io.StringIO(decode(z.read(n)))); h=next(rdr)
        si=cidx(h,"시리얼"); ni=cidx(h,"소음"); ri=cidx(h,"등록")
        for r in rdr:
            if si<len(r) and ni<len(r):
                v=r[ni].strip()
                if not v: continue
                try: fv=float(v)
                except: continue
                if fv<=0: continue
                lin=10**(fv/10.0)
                ts=r[ri].strip() if ri<len(r) else ""
                mm=0; hh=None
                if ts:
                    dp=ts.split(); dpart=dp[0].replace(".","-").split("-")
                    if len(dpart)>=2 and dpart[1].isdigit(): mm=int(dpart[1])
                    if len(dp)>1:
                        try: hh=int(dp[1].split(":")[0])
                        except: hh=None
                s=r[si].strip(); q=(mm-1)//3+1 if 1<=mm<=12 else 0
                acc[s]["all"][0]+=lin; acc[s]["all"][1]+=1
                if q: acc[s][f"Q{q}"][0]+=lin; acc[s][f"Q{q}"][1]+=1
                if hh is not None:
                    k='day' if 6<=hh<=21 else 'night'
                    acc[s][k][0]+=lin; acc[s][k][1]+=1
    def leq(p): return round(10*math.log10(p[0]/p[1]),2) if p[1]>0 else ""
    with open(os.path.join(sdot_dir,outcsv),"w",newline="",encoding="utf-8-sig") as f:
        w=csv.writer(f)
        w.writerow(["serial","lat","lon","n_all","Leq_annual","Leq_Q1","Leq_Q2","Leq_Q3","Leq_Q4","Leq_day","Leq_night"])
        for s,d in sorted(acc.items()):
            c=ser2c.get(s,("",""))
            w.writerow([s,c[0],c[1],d["all"][1],leq(d["all"]),leq(d.get("Q1",[0,0])),leq(d.get("Q2",[0,0])),
                        leq(d.get("Q3",[0,0])),leq(d.get("Q4",[0,0])),leq(d.get("day",[0,0])),leq(d.get("night",[0,0]))])
    matched=sum(1 for s in acc if s in ser2c)
    print(f"{zname} -> {outcsv}: sensors={len(acc)}, coord-matched={matched}")

build("sdot2022.zip","sdot_levels_2022.csv")
build("sdot2020.zip","sdot_levels_2020.csv")
print("=== DONE ===")
