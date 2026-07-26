# -*- coding: utf-8 -*-
# Phase 2-C 변화량 공간패턴 (파스텔·Arial·지리비율) — 2x2 layout(본문 단 너비에 적합한 큰 panel):
#  (a) 토지이용 분류  (b) 주간 이동량 감소율  (c) 드리프트제거 상대 소음변화  (d) 이동량-소음 산점도(토지이용별).
import os, json
import numpy as np
import pandas as pd
from scipy import stats
import matplotlib.pyplot as plt
from matplotlib.collections import PolyCollection
from matplotlib.cm import ScalarMappable
from matplotlib.colors import Normalize
from matplotlib.patches import Patch
import figstyle as FS

FS.apply_style()
ROOT = r"T:\00_BACKUP\01_한양대학교\00_연구실\00_프로젝트\★_논문화 프로젝트\31_환경소음_이동성"
P = os.path.join(ROOT, "data", "processed")
FIG = os.path.join(ROOT, "_claude", "figures")
GEO = os.path.join(ROOT, "data", "reference", "admdong_seoul_ver20220101.geojson")
MIN_SENS = 2   # 추세·상관은 신뢰도 위해 센서 2개 이상 동만 (단일센서 동은 그 센서 드리프트가 곧 동값)

df = pd.read_csv(os.path.join(P, "analysis_panel.csv"), dtype={"serial": str, "dose_key": str},
                 usecols=["serial", "dose_key", "date", "dLeq_day", "lp_day_rel", "lp_day", "lp_night", "stringency"])
df["date"] = pd.to_datetime(df["date"])
post = df[(df["date"] >= "2022-07-01") & (df["date"] <= "2023-12-31")]
comm = post.groupby("dose_key").apply(lambda g: g["lp_day"].mean() / g["lp_night"].mean(),
                                      include_groups=False).rename("comm_idx")
rest = df[df["stringency"] >= 4]
city = rest["dLeq_day"].mean()
g = rest.groupby("dose_key").agg(dLeq=("dLeq_day", "mean"), mob=("lp_day_rel", "mean"),
                                 n=("dLeq_day", "size"), n_sens=("serial", "nunique")).reset_index().merge(comm, on="dose_key")
g["noise_anom"] = g["dLeq"] - city
g["mob_drop"] = (1 - g["mob"]) * 100
g = g[g["n"] >= 100].copy()
q1, q2 = g["comm_idx"].quantile([1/3, 2/3])
g["landuse"] = np.where(g["comm_idx"] >= q2, "commercial",
                        np.where(g["comm_idx"] <= q1, "residential", "mixed"))
stat = g.set_index("dose_key")
print(f"지도 동 {len(g)} | 상업 {(g.landuse=='commercial').sum()} 혼합 {(g.landuse=='mixed').sum()} 주거 {(g.landuse=='residential').sum()}")

REV = {"11740525": "11740520", "11740526": "11740520"}
geo = json.load(open(GEO, encoding="utf-8"))
def polys_of(gm):
    return gm["coordinates"] if gm["type"] == "MultiPolygon" else [gm["coordinates"]]
PATS, KEYS = [], []
for f in geo["features"]:
    code = str(f["properties"]["adm_cd2"])[:8]
    key = REV.get(code, code)
    for poly in polys_of(f["geometry"]):
        PATS.append(np.array(poly[0])); KEYS.append(key)

def draw_cont(ax, col, cmap, vlim, title, clabel):
    norm = Normalize(*vlim); cm = plt.get_cmap(cmap)
    cols = [cm(norm(stat.loc[k, col])) if k in stat.index else FS.NA_FILL for k in KEYS]
    ax.add_collection(PolyCollection(PATS, facecolors=cols, edgecolors=FS.EDGE, linewidths=.25))
    ax.autoscale(); FS.style_map_ax(ax); ax.set_title(title)
    sm = ScalarMappable(norm=norm, cmap=cmap); sm.set_array([])
    cb = plt.colorbar(sm, ax=ax, fraction=.043, pad=.02); cb.set_label(clabel, fontsize=9); cb.outline.set_visible(False)

def draw_cat(ax, title):
    cm = {k: FS.LANDUSE_COLORS[stat.loc[k, "landuse"]] for k in stat.index}
    cols = [cm.get(k, FS.NA_FILL) for k in KEYS]
    ax.add_collection(PolyCollection(PATS, facecolors=cols, edgecolors=FS.EDGE, linewidths=.25))
    ax.autoscale(); FS.style_map_ax(ax); ax.set_title(title)
    SHORT = {"commercial": "Commercial", "mixed": "Mixed", "residential": "Residential"}
    ax.legend(handles=[Patch(facecolor=FS.LANDUSE_COLORS[k], edgecolor="white", label=SHORT[k])
                       for k in ["commercial", "mixed", "residential"]],
              loc="upper center", bbox_to_anchor=(0.5, -0.02), ncol=3, fontsize=8,
              handlelength=1.1, columnspacing=1.4, frameon=False)   # 지도 아래로 빼 겹침 방지

def draw_scatter(ax):
    # 신뢰도 위해 센서>=MIN_SENS 동만, robust(Theil-Sen)+Spearman, 점 크기=센서수
    gg = g[g["n_sens"] >= MIN_SENS].dropna(subset=["mob_drop", "noise_anom"])
    for lu in ["residential", "mixed", "commercial"]:
        d = gg[gg["landuse"] == lu]
        c = FS.LANDUSE_COLORS[lu]
        rho = stats.spearmanr(d["mob_drop"], d["noise_anom"])[0] if len(d) > 5 else float("nan")
        ax.scatter(d["mob_drop"], d["noise_anom"], s=8 + d["n_sens"] * 5, color=c, alpha=.7,
                   edgecolors="white", linewidths=.3,
                   label=f"{FS.LANDUSE_LABEL[lu].split(' (')[0]} (ρ={rho:+.2f}, n={len(d)})")
        if len(d) > 5 and d["mob_drop"].std() > 1e-6:
            ts = stats.theilslopes(d["noise_anom"].values, d["mob_drop"].values)
            xs = np.linspace(d["mob_drop"].min(), d["mob_drop"].max(), 50)
            ax.plot(xs, ts[1] + ts[0] * xs, color=c, lw=1.8)
    ax.axhline(0, color=FS.ACCENT["neutral"], lw=.7, ls="--")
    ax.set_xlabel("Daytime mobility reduction (%)")
    ax.set_ylabel("Relative noise change ΔL$_{day}$ (dB)")
    ax.set_title("(d)", fontsize=11)
    ax.set_xlim(-20, 20); ax.set_ylim(-5, 5)   # 이상치(소수 동) 제외하고 본체에 맞춰 — 회귀·상관은 전체 데이터 기준
    ax.legend(loc="upper right", fontsize=7.4)

fig, ax = plt.subplots(2, 2, figsize=(11.6, 10.2))
draw_cat(ax[0, 0], "(a)")
draw_cont(ax[0, 1], "mob_drop", FS.PASTEL_SEQ, (0, max(8, g["mob_drop"].quantile(.95))),
          "(b)", "Mobility drop (%)")
nlim = max(abs(g["noise_anom"].quantile(.05)), abs(g["noise_anom"].quantile(.95)))
draw_cont(ax[1, 0], "noise_anom", FS.PASTEL_DIV, (-nlim, nlim),
          "(c)", "ΔL$_{day}$ vs city mean (dB)")
draw_scatter(ax[1, 1])
gg = g[g["n_sens"] >= MIN_SENS].dropna(subset=["mob_drop", "noise_anom"])
rho = stats.spearmanr(gg["mob_drop"], gg["noise_anom"])[0]
# suptitle 제거(캡션이 대신함)
plt.tight_layout(rect=[0, 0, 1, 0.98])
out = os.path.join(FIG, "fig_change_map.png")
plt.savefig(out)
print(f"robust 공간상관 Spearman ρ(이동량감소율 vs 상대소음변화, n_sens>={MIN_SENS}, n={len(gg)}) = {rho:+.3f}")
print(f"-> {out}")
g.to_csv(os.path.join(P, "phase2c_dong_change.csv"), index=False, encoding="utf-8-sig")
