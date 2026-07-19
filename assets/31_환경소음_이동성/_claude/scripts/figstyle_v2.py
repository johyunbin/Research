# -*- coding: utf-8 -*-
# 그림 스타일 v2 — SCI 저널용 고대비·colorblind-safe(Okabe-Ito) 체계.
# 의미 색 고정: mobility=blue, noise=vermilion, 토지이용 3색(상업 bluish-green/혼합 grey/주거 orange).
# 지도 공용 헬퍼(폴리곤·스케일바·방위·패널라벨)와 거리두기 음영 포함. 모든 figs_v2_*에서 import.
import os, json, math
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.collections import PolyCollection

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
P = os.path.join(ROOT, "data", "processed")
REF = os.path.join(ROOT, "data", "reference")
FIG = os.path.join(ROOT, "_claude", "figures_v2")
CACHE = os.path.join(FIG, "cache")
os.makedirs(FIG, exist_ok=True)
os.makedirs(CACHE, exist_ok=True)

def apply_style():
    plt.rcParams.update({
        "font.family": "Arial",
        "font.sans-serif": ["Arial", "DejaVu Sans"],
        "font.size": 9,
        "axes.titlesize": 10,
        "axes.titleweight": "bold",
        "axes.labelsize": 9.5,
        "axes.edgecolor": "#333333",
        "axes.linewidth": 0.9,
        "axes.spines.top": False,
        "axes.spines.right": False,
        "xtick.labelsize": 8.5,
        "ytick.labelsize": 8.5,
        "xtick.color": "#333333",
        "ytick.color": "#333333",
        "xtick.direction": "out",
        "ytick.direction": "out",
        "legend.fontsize": 8.5,
        "legend.frameon": False,
        "figure.dpi": 110,
        "savefig.dpi": 400,
        "savefig.bbox": "tight",
        "savefig.facecolor": "white",
        "axes.grid": False,
        "pdf.fonttype": 42,
    })

# ---- Okabe-Ito 기반 의미 색 ----
MOB = "#0072B2"       # mobility (blue)
NOISE = "#D55E00"     # noise (vermilion)
NEUTRAL = "#5A5A5A"
SIG = "#0072B2"       # FDR 유의
NS = "#ABABAB"        # 비유의
LANDUSE = {"commercial": "#009E73", "mixed": "#8C8C8C", "residential": "#E69F00"}
LANDUSE_LABEL = {"commercial": "Commercial", "mixed": "Mixed", "residential": "Residential"}
NA_FILL = "#F0F0EE"
EDGE = "white"

# 지도 컬러맵(모든 지도 공통 의미: red=많음/시끄러움, blue=적음/조용)
CMAP_DIV = "RdBu_r"
CMAP_SEQ_NOISE = "viridis"
CMAP_SEQ_DROP = "Blues"

SEOUL_LAT = 37.565
def map_aspect():
    return 1.0 / math.cos(math.radians(SEOUL_LAT))

def style_map_ax(ax):
    ax.set_aspect(map_aspect())
    ax.axis("off")

# ---- 서울 행정동 폴리곤 (crosswalk 포함) ----
REV = {"11740525": "11740520", "11740526": "11740520"}
def load_seoul():
    geo = json.load(open(os.path.join(REF, "admdong_seoul_ver20220101.geojson"), encoding="utf-8"))
    pats, keys = [], []
    for f in geo["features"]:
        code = str(f["properties"]["adm_cd2"])[:8]
        key = REV.get(code, code)
        gm = f["geometry"]
        polys = gm["coordinates"] if gm["type"] == "MultiPolygon" else [gm["coordinates"]]
        for poly in polys:
            pats.append(np.array(poly[0])); keys.append(key)
    return pats, keys

def draw_polys(ax, pats, colors, lw=0.3):
    ax.add_collection(PolyCollection(pats, facecolors=colors, edgecolors=EDGE, linewidths=lw))
    ax.autoscale()
    style_map_ax(ax)

# ---- 지도 부속: 스케일바·방위 ----
def add_scalebar(ax, km=5, loc=(0.72, 0.045)):
    deg = km / (111.32 * math.cos(math.radians(SEOUL_LAT)))
    x0, x1 = ax.get_xlim(); y0, y1 = ax.get_ylim()
    xs = x0 + loc[0] * (x1 - x0); ys = y0 + loc[1] * (y1 - y0)
    ax.plot([xs, xs + deg], [ys, ys], color="#222222", lw=2.2, solid_capstyle="butt", clip_on=False)
    ax.text(xs + deg / 2, ys + 0.012 * (y1 - y0), f"{km} km", ha="center", va="bottom",
            fontsize=7.5, color="#222222")

def add_north(ax, loc=(0.955, 0.90)):
    x0, x1 = ax.get_xlim(); y0, y1 = ax.get_ylim()
    xs = x0 + loc[0] * (x1 - x0); ys = y0 + loc[1] * (y1 - y0)
    dy = 0.055 * (y1 - y0)
    ax.annotate("", xy=(xs, ys + dy), xytext=(xs, ys),
                arrowprops=dict(arrowstyle="-|>", color="#222222", lw=1.4))
    ax.text(xs, ys + dy * 1.22, "N", ha="center", va="bottom", fontsize=8.5,
            fontweight="bold", color="#222222")

# ---- 패널 라벨 ----
def panel_label(ax, s, dx=0.0, dy=0.0):
    ax.text(0.0 + dx, 1.02 + dy, f"({s})", transform=ax.transAxes, fontsize=11,
            fontweight="bold", va="bottom", ha="left")

# ---- 거리두기 주요 전환점·강제한 음영 ----
EVENTS = [("2020-08-30", "21:00 curfew"), ("2020-12-23", "5-person ban"),
          ("2021-07-12", "Level 4"), ("2021-11-01", "With-COVID"),
          ("2021-12-18", "Re-tighten"), ("2022-04-18", "Full lifting")]

def restriction_windows(min_stringency=4):
    d = pd.read_csv(os.path.join(P, "distancing_daily_2020-2023.csv"))
    dcol = "date" if "date" in d.columns else d.columns[0]
    scol = "stringency" if "stringency" in d.columns else d.columns[-1]
    d[dcol] = pd.to_datetime(d[dcol])
    d = d.sort_values(dcol)
    on = d[scol] >= min_stringency
    wins, start = [], None
    for dt, flag in zip(d[dcol], on):
        if flag and start is None:
            start = dt
        elif not flag and start is not None:
            wins.append((start, dt)); start = None
    if start is not None:
        wins.append((start, d[dcol].iloc[-1]))
    return wins

def shade_restrictions(ax, min_stringency=4, color="#B0B0B0", alpha=0.14):
    for s, e in restriction_windows(min_stringency):
        ax.axvspan(s, e, color=color, alpha=alpha, lw=0, zorder=0)

def mark_events(ax, y_frac=1.0, fontsize=6.8, labels=True):
    for dstr, lab in EVENTS:
        ax.axvline(pd.Timestamp(dstr), color=NEUTRAL, lw=0.6, alpha=0.45, zorder=1)
    if labels:
        for dstr, lab in EVENTS:
            ax.annotate(lab, (pd.Timestamp(dstr), y_frac), xycoords=("data", "axes fraction"),
                        xytext=(2, -1), textcoords="offset points", fontsize=fontsize,
                        rotation=90, va="top", ha="left", color="#555555")
