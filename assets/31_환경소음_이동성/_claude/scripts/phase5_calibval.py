# -*- coding: utf-8 -*-
# Phase 5 Tier1: 검교정 환경소음망(서울 146점, 2024 LAeq) vs 인근 S-DoT 레벨(2022) — 절대레벨 bias 정량 + 공간 타당성.
# 목적: S-DoT 절대레벨이 검교정 LAeq와 (a)얼마나 어긋나는가(offset) (b)공간적으로 같은 패턴을 잡는가(상관).
import os, csv, math
import statistics as st
import matplotlib.pyplot as plt
import figstyle as FS

FS.apply_style()
ROOT = r"T:\00_BACKUP\01_한양대학교\00_연구실\00_프로젝트\★_논문화 프로젝트\31_환경소음_이동성"
P = os.path.join(ROOT, "data", "processed")
FIG = os.path.join(ROOT, "_claude", "figures")
def rd(p): return list(csv.DictReader(open(p, encoding="utf-8-sig")))
def hav(a1, o1, a2, o2):
    R = 6371000.0; la1, lo1, la2, lo2 = map(math.radians, (a1, o1, a2, o2))
    h = math.sin((la2-la1)/2)**2 + math.cos(la1)*math.cos(la2)*math.sin((lo2-lo1)/2)**2
    return 2*R*math.asin(math.sqrt(h))
def emean(vs):
    vs = [v for v in vs if v is not None]
    return 10*math.log10(sum(10**(v/10) for v in vs)/len(vs)) if vs else None
def corr(xs, ys, rank=False):
    n = len(xs)
    if n < 8: return None
    if rank:
        def rk(v):
            o = sorted(range(len(v)), key=lambda i: v[i]); r = [0]*len(v)
            for pos, i in enumerate(o): r[i] = pos
            return r
        xs, ys = rk(xs), rk(ys)
    mx = sum(xs)/n; my = sum(ys)/n
    sxx = sum((x-mx)**2 for x in xs); syy = sum((y-my)**2 for y in ys); sxy = sum((x-mx)*(y-my) for x, y in zip(xs, ys))
    return sxy/math.sqrt(sxx*syy) if sxx > 0 and syy > 0 else None

# 검교정 146점: 좌표 + 2024 LAeq 결합
co = {r["측정지점"]: r for r in rd(os.path.join(P, "ref146_coords.csv"))}
ref = []
for r in rd(os.path.join(P, "reference_seoul_146.csv")):
    c = co.get(r["측정지점"])
    if not c or not c["위도"] or not c["경도"]: continue
    try:
        ref.append({"name": r["측정지점"], "zone": r["지역"], "lat": float(c["위도"]), "lon": float(c["경도"]),
                    "cd": float(r["Leq_day_2024"]), "cn": float(r["Leq_night_2024"])})
    except: continue
print(f"검교정 점(좌표+값 유효): {len(ref)} / 146")

# S-DoT 2022 연레벨
sd = []
for r in rd(os.path.join(P, "sdot_levels_2022.csv")):
    try:
        lat, lon, n = float(r["lat"]), float(r["lon"]), int(r["n_all"])
        ld = float(r["Leq_day"]) if r["Leq_day"] else None
        ln = float(r["Leq_night"]) if r["Leq_night"] else None
    except: continue
    if n < 2000 or ld is None: continue
    sd.append((lat, lon, ld, ln))
print(f"S-DoT QC 센서(2022, n≥2000): {len(sd)}")

# 매칭: 각 검교정점 R 내 S-DoT 에너지평균
for R in (300.0, 500.0):
    rows = []
    for q in ref:
        ds = [(hav(q["lat"], q["lon"], a, o), ld, ln) for a, o, ld, ln in sd]
        near = [(ld, ln) for d, ld, ln in ds if d <= R]
        if len(near) < 1: continue
        sdd = emean([x[0] for x in near]); sdn = emean([x[1] for x in near])
        rows.append({**q, "sdd": sdd, "sdn": sdn, "k": len(near),
                     "offd": sdd - q["cd"], "offn": (sdn - q["cn"]) if sdn else None})
    nd = [r["offd"] for r in rows]
    nn = [r["offn"] for r in rows if r["offn"] is not None]
    pr = corr([r["cd"] for r in rows], [r["sdd"] for r in rows])
    sp = corr([r["cd"] for r in rows], [r["sdd"] for r in rows], rank=True)
    print(f"\n=== R={R:.0f}m | 매칭 {len(rows)}점 (평균 {st.mean(r['k'] for r in rows):.1f} 센서/점) ===")
    print(f"  주간 offset(S-DoT−LAeq): 평균 {st.mean(nd):+.2f} dB, 중앙 {st.median(nd):+.2f}, SD {st.pstdev(nd):.2f}")
    print(f"  야간 offset: 평균 {st.mean(nn):+.2f} dB, 중앙 {st.median(nn):+.2f}")
    print(f"  공간상관(검교정 LAeq vs 인근 S-DoT, 주간): Pearson {pr:+.2f} · Spearman {sp:+.2f}")
    # 지역별
    for z in ("일반", "도로"):
        zr = [r["offd"] for r in rows if r["zone"] == z]
        if zr: print(f"    [{z}] n={len(zr)} offset 평균 {st.mean(zr):+.2f} dB")
    if R == 500.0:
        save_rows, save_pr, save_sp = rows, pr, sp

# 저장 + 그림(R=500)
with open(os.path.join(P, "phase5_calibval.csv"), "w", encoding="utf-8-sig", newline="") as f:
    w = csv.writer(f); w.writerow(["name", "zone", "k", "cal_day", "sdot_day", "offset_day", "cal_night", "sdot_night"])
    for r in save_rows: w.writerow([r["name"], r["zone"], r["k"], f"{r['cd']:.1f}", f"{r['sdd']:.1f}", f"{r['offd']:+.1f}", f"{r['cn']:.1f}", f"{r['sdn']:.1f}" if r['sdn'] else ""])

fig, ax = plt.subplots(1, 2, figsize=(10.5, 5.0))
ZC = {"일반": FS.ACCENT["mobility"], "도로": FS.ACCENT["noise"]}
for z in ("일반", "도로"):
    xs = [r["cd"] for r in save_rows if r["zone"] == z]; ys = [r["sdd"] for r in save_rows if r["zone"] == z]
    ax[0].scatter(xs, ys, s=34, color=ZC[z], alpha=.8, edgecolors="white", linewidths=.5,
                  label={"일반": "Residential/general", "도로": "Roadside"}[z])
lo, hi = 40, 80
ax[0].plot([lo, hi], [lo, hi], ls=":", color="#bbb", lw=1, zorder=0)
ax[0].set_xlim(lo, hi); ax[0].set_ylim(lo, hi); ax[0].set_box_aspect(1)
ax[0].set_xlabel("Calibrated L$_{Aeq}$ daytime, 2024 (dB)")
ax[0].set_ylabel("Nearby S-DoT L$_{day}$, 2022 (dB)")
ax[0].set_title("(a)"); ax[0].legend(loc="upper left", fontsize=8.5)
ax[0].text(.04, .70, f"Spearman ρ={save_sp:+.2f}\nPearson r={save_pr:+.2f}\nn={len(save_rows)} (≤500 m)",
           transform=ax[0].transAxes, fontsize=8.5, bbox=dict(boxstyle="round,pad=0.3", fc="white", ec="#bbb"))
offs = [r["offd"] for r in save_rows]
ax[1].hist(offs, bins=18, color="#C9D6C3", edgecolor="#7a8a72")
ax[1].axvline(st.mean(offs), color=FS.ACCENT["noise"], lw=2, label=f"mean {st.mean(offs):+.1f} dB")
ax[1].axvline(0, color=FS.ACCENT["neutral"], lw=.8, ls="--")
ax[1].set_xlabel("S-DoT − calibrated L$_{Aeq}$ offset, daytime (dB)"); ax[1].set_ylabel("Stations")
ax[1].set_title("(b)"); ax[1].legend(loc="upper right", fontsize=8.5); ax[1].set_box_aspect(1)
plt.tight_layout()
out = os.path.join(FIG, "fig_calibval.png")
plt.savefig(out)
print(f"\n-> {out}\n-> phase5_calibval.csv")
