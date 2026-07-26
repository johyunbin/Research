# -*- coding: utf-8 -*-
# 내부 공간정합성: 인접 S-DoT 센서 레벨 일치(공간 자기상관) — 전 지역 신뢰성 검증
import csv, math, os, statistics
from collections import defaultdict

sdot_dir = r"C:\Users\wh850\AppData\Local\Temp\sdot"
rows = list(csv.DictReader(open(os.path.join(sdot_dir,"sdot_levels_2022.csv"), encoding="utf-8-sig")))

# QC 필터: 물리적 레벨 + 최소 관측수
pts=[]; dropped=0
for r in rows:
    try:
        la=float(r["Leq_annual"]); lat=float(r["lat"]); lon=float(r["lon"]); n=int(r["n_all"])
    except: dropped+=1; continue
    if la<35 or la>85 or n<2000: dropped+=1; continue
    pts.append((lat,lon,la))
print(f"QC: kept {len(pts)} sensors, dropped {dropped} (level<35/>85 or n<2000)")
allv=[p[2] for p in pts]
print(f"Leq_annual after QC: mean {statistics.mean(allv):.1f}, sd {statistics.pstdev(allv):.1f} (= '실(sill)' 기준)")

def hav(a1,o1,a2,o2):
    R=6371000.0; la1,lo1,la2,lo2=map(math.radians,(a1,o1,a2,o2))
    h=math.sin((la2-la1)/2)**2+math.cos(la1)*math.cos(la2)*math.sin((lo2-lo1)/2)**2
    return 2*R*math.asin(math.sqrt(h))

bins=[(0,100),(100,200),(200,300),(300,500),(500,1000),(1000,2000)]
acc={b:[] for b in bins}
nn_self=[]; nn_neigh=[]   # 최근접이웃 쌍(<=300m) 레벨 상관용
n=len(pts)
for i in range(n):
    a1,o1,l1=pts[i]; best=None
    for j in range(i+1,n):
        a2,o2,l2=pts[j]
        if abs(a1-a2)>0.02 or abs(o1-o2)>0.025: continue
        d=hav(a1,o1,a2,o2)
        if d>=2000: continue
        for b in bins:
            if b[0]<=d<b[1]: acc[b].append(abs(l1-l2)); break
        if d<=300 and (best is None or d<best[0]): best=(d,l2)
    if best: nn_self.append(l1); nn_neigh.append(best[1])

print("\n[공간 semivariogram] 거리bin | 쌍수 | median|ΔLeq| | mean|ΔLeq|")
for b in bins:
    v=acc[b]
    if v: print(f"  {b[0]:4d}-{b[1]:4d}m | {len(v):6d} | {statistics.median(v):4.1f} | {statistics.mean(v):4.1f} dB")

def pear(xs,ys):
    n=len(xs); mx=sum(xs)/n; my=sum(ys)/n
    sxx=sum((x-mx)**2 for x in xs); syy=sum((y-my)**2 for y in ys); sxy=sum((x-mx)*(y-my) for x,y in zip(xs,ys))
    return sxy/math.sqrt(sxx*syy) if sxx>0 and syy>0 else None
if len(nn_self)>=20:
    r=pear(nn_self,nn_neigh)
    print(f"\n[최근접이웃(<=300m) 레벨 상관] n_pairs={len(nn_self)}  Pearson r={r:.3f}")
    diffs=[abs(a-b) for a,b in zip(nn_self,nn_neigh)]
    print(f"  최근접이웃 |ΔLeq|: median {statistics.median(diffs):.1f} dB, 90퍼센타일 {sorted(diffs)[int(len(diffs)*0.9)]:.1f} dB")
print("\n해석 가이드: 단거리 |ΔLeq|가 작고 거리 따라 증가 + 최근접 상관 높으면 → 공간 노이즈장이 실재(센서 신뢰).")
print("=== DONE ===")
