# -*- coding: utf-8 -*-
# 외부 다지점 공간검증: 기준 LAeq(146점, 2024) vs 인근 S-DoT 레벨(2022) — 좌표 도착 후 실행
import csv, math, os, statistics
sdot_dir = r"C:\Users\wh850\AppData\Local\Temp\sdot"
def rd(p): return list(csv.DictReader(open(p, encoding="utf-8-sig")))
def hav(a1,o1,a2,o2):
    R=6371000.0; la1,lo1,la2,lo2=map(math.radians,(a1,o1,a2,o2))
    h=math.sin((la2-la1)/2)**2+math.cos(la1)*math.cos(la2)*math.sin((lo2-lo1)/2)**2
    return 2*R*math.asin(math.sqrt(h))
def pear(xs,ys):
    n=len(xs)
    if n<8: return None
    mx=sum(xs)/n; my=sum(ys)/n
    sxx=sum((x-mx)**2 for x in xs); syy=sum((y-my)**2 for y in ys); sxy=sum((x-mx)*(y-my) for x,y in zip(xs,ys))
    return sxy/math.sqrt(sxx*syy) if sxx>0 and syy>0 else None
def spear(xs,ys):
    def rk(v):
        o=sorted(range(len(v)),key=lambda i:v[i]); r=[0]*len(v)
        for pos,i in enumerate(o): r[i]=pos
        return r
    return pear(rk(xs),rk(ys))

coords_p=os.path.join(sdot_dir,"ref146_coords.csv")
if not os.path.exists(coords_p):
    print("ref146_coords.csv 아직 없음 — 지오코딩 에이전트 완료 후 실행"); raise SystemExit
coords=rd(coords_p)
refval={r["측정지점"]:r for r in rd(os.path.join(sdot_dir,"reference_seoul_146.csv"))}
sd=[]
for r in rd(os.path.join(sdot_dir,"sdot_levels_2022.csv")):
    try:
        la=float(r["Leq_annual"]); lat=float(r["lat"]); lon=float(r["lon"]); n=int(r["n_all"])
        ld=float(r["Leq_day"]) if r["Leq_day"] else None
    except: continue
    if la<35 or la>85 or n<2000: continue
    sd.append((lat,lon,la,ld))
print("S-DoT QC sensors:", len(sd))

def cval(r):
    try: return float(r["위도"]), float(r["경도"])
    except: return None
maxR=500.0
pairs=[]
for r in coords:
    cv=cval(r)
    if not cv: continue
    pt=r.get("측정지점","").strip(); rv=refval.get(pt)
    if not rv or not rv.get("Leq_day_2024"): continue
    try: rlaeq=float(rv["Leq_day_2024"])
    except: continue
    best=None; clu=[]
    for lat,lon,la,ld in sd:
        d=hav(cv[0],cv[1],lat,lon)
        if d<=maxR: clu.append(ld)
        if best is None or d<best[0]: best=(d,la,ld)
    if best and best[0]<=maxR:
        days=[x for x in clu if x is not None]
        clu_day=10*math.log10(sum(10**(x/10) for x in days)/len(days)) if days else None
        pairs.append((pt,rlaeq,best[1],best[2],best[0],r.get("confidence","").strip(),clu_day))

print(f"matched ref points (<= {maxR:.0f}m): {len(pairs)} / coords rows {len(coords)}")
if len(pairs)>=8:
    ad=[(p[1],p[3]) for p in pairs if p[3] is not None]
    an=[(p[1],p[2]) for p in pairs if p[2] is not None]
    ac=[(p[1],p[6]) for p in pairs if p[6] is not None]
    print(f"[ref dayLAeq vs S-DoT day]    n={len(ad)} Pearson={pear([a for a,_ in ad],[b for _,b in ad]):.3f} Spearman={spear([a for a,_ in ad],[b for _,b in ad]):.3f}")
    print(f"[ref dayLAeq vs S-DoT annual] n={len(an)} Pearson={pear([a for a,_ in an],[b for _,b in an]):.3f}")
    print(f"[ref dayLAeq vs cluster<=R]   n={len(ac)} Pearson={pear([a for a,_ in ac],[b for _,b in ac]):.3f}")
    hi=[p for p in pairs if p[5] in ("high","authoritative") and p[3] is not None]
    if len(hi)>=8:
        print(f"[high-confidence only]        n={len(hi)} Pearson={pear([p[1] for p in hi],[p[3] for p in hi]):.3f}")
    bias=statistics.mean([p[1]-p[3] for p in pairs if p[3] is not None])
    print(f"mean bias (ref - S-DoT day) = {bias:+.1f} dB")
else:
    print("matched points < 8 — 좌표 확보 부족")
print("=== DONE ===")
