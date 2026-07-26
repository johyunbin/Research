# -*- coding: utf-8 -*-
# 공간검증 near-zero 원인 진단: 매칭쌍 덤프 + 근접도/신뢰도 부분집합 상관
import csv, math, os, statistics
sdot_dir=r"C:\Users\wh850\AppData\Local\Temp\sdot"
def rd(p): return list(csv.DictReader(open(p,encoding="utf-8-sig")))
def hav(a1,o1,a2,o2):
    R=6371000.0; la1,lo1,la2,lo2=map(math.radians,(a1,o1,a2,o2))
    h=math.sin((la2-la1)/2)**2+math.cos(la1)*math.cos(la2)*math.sin((lo2-lo1)/2)**2
    return 2*R*math.asin(math.sqrt(h))
def pear(xs,ys):
    n=len(xs)
    if n<6: return None
    mx=sum(xs)/n; my=sum(ys)/n
    sxx=sum((x-mx)**2 for x in xs); syy=sum((y-my)**2 for y in ys); sxy=sum((x-mx)*(y-my) for x,y in zip(xs,ys))
    return sxy/math.sqrt(sxx*syy) if sxx>0 and syy>0 else None

coords=rd(os.path.join(sdot_dir,"ref146_coords.csv"))
refval={r["측정지점"]:r for r in rd(os.path.join(sdot_dir,"reference_seoul_146.csv"))}
sd=[]
for r in rd(os.path.join(sdot_dir,"sdot_levels_2022.csv")):
    try:
        la=float(r["Leq_annual"]); lat=float(r["lat"]); lon=float(r["lon"]); n=int(r["n_all"])
        ld=float(r["Leq_day"]) if r["Leq_day"] else None
    except: continue
    if la<35 or la>85 or n<2000: continue
    sd.append((lat,lon,la,ld))

pairs=[]
for r in coords:
    try: lat0,lon0=float(r["위도"]),float(r["경도"])
    except: continue
    pt=r.get("측정지점","").strip(); rv=refval.get(pt)
    if not rv or not rv.get("Leq_day_2024"): continue
    try: rl=float(rv["Leq_day_2024"])
    except: continue
    best=None
    for lat,lon,la,ld in sd:
        d=hav(lat0,lon0,lat,lon)
        if best is None or d<best[0]: best=(d,la,ld)
    if best and best[2] is not None:
        pairs.append((pt,rl,best[2],best[0],r.get("confidence","").strip()))

print("ref point | refDayLAeq | sdotDay | dist_m | conf")
for p in sorted(pairs,key=lambda x:-x[1])[:30]:
    print(f"  {p[0][:22]:22s} | {p[1]:5.1f} | {p[2]:5.1f} | {p[3]:5.0f} | {p[4]}")
print(f"  ... ({len(pairs)} total)")

print("\nsubset | n | Pearson")
def sub(name, flt):
    pp=[p for p in pairs if flt(p)]
    r=pear([p[1] for p in pp],[p[2] for p in pp])
    print(f"  {name:28s} | {len(pp):3d} | {('%.3f'%r) if r is not None else 'na'}")
sub("all (<=nearest)", lambda p: True)
sub("dist<=300m", lambda p: p[3]<=300)
sub("dist<=200m", lambda p: p[3]<=200)
sub("dist<=150m", lambda p: p[3]<=150)
sub("dist<=100m", lambda p: p[3]<=100)
sub("conf in {high,authoritative}", lambda p: p[4] in ("high","authoritative"))
sub("high/auth & dist<=200m", lambda p: p[4] in ("high","authoritative") and p[3]<=200)
print(f"\nref range: {min(p[1] for p in pairs):.0f}-{max(p[1] for p in pairs):.0f} (sd {statistics.pstdev([p[1] for p in pairs]):.1f})")
print(f"sdot range: {min(p[2] for p in pairs):.0f}-{max(p[2] for p in pairs):.0f} (sd {statistics.pstdev([p[2] for p in pairs]):.1f})")
print("=== DONE ===")
