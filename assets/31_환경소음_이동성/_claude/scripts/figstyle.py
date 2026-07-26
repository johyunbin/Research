# -*- coding: utf-8 -*-
# 공유 그림 스타일: Arial · 파스텔 컬러맵 · 서울 지도 종횡비. 모든 phase2 그림에서 import.
import math
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.colors import LinearSegmentedColormap

def apply_style():
    plt.rcParams.update({
        "font.family": "Arial",
        "font.sans-serif": ["Arial", "DejaVu Sans"],
        "font.size": 10,
        "axes.titlesize": 11,
        "axes.titleweight": "bold",
        "axes.labelsize": 10,
        "axes.edgecolor": "#888888",
        "axes.linewidth": 0.8,
        "xtick.labelsize": 8.5,
        "ytick.labelsize": 8.5,
        "legend.fontsize": 8.5,
        "legend.frameon": False,
        "figure.dpi": 110,
        "savefig.dpi": 300,
        "savefig.bbox": "tight",
        "axes.grid": False,
    })

# --- 파스텔 컬러맵 ---
# 순차(이동량 감소 등 0->high): cream -> soft peach -> muted coral
PASTEL_SEQ = LinearSegmentedColormap.from_list(
    "pastel_seq", ["#FBF6EF", "#F6DEC4", "#EFBE9C", "#E09B82"])
# 발산(소음 anomaly -/0/+): muted teal <- cream -> muted rose
PASTEL_DIV = LinearSegmentedColormap.from_list(
    "pastel_div", ["#6FA8B5", "#BFD8D6", "#F2EDE4", "#E7C4BD", "#CE8C86"])
# 발산 역(필요시)
PASTEL_DIV_R = PASTEL_DIV.reversed()

# 토지이용 3분류 파스텔 (상업/혼합/주거)
LANDUSE_COLORS = {"commercial": "#7FB8A4", "mixed": "#DDD3C2", "residential": "#A9A4C9"}
LANDUSE_LABEL = {"commercial": "Commercial (high day/night ratio)",
                 "mixed": "Mixed", "residential": "Residential (low ratio)"}
ACCENT = {"noise": "#C0708A", "mobility": "#5FA88C", "neutral": "#7B8794"}

NA_FILL = "#EFEFEC"   # 데이터 없는 동
EDGE = "white"

SEOUL_LAT = 37.565
def map_aspect():
    """경위도 지도를 실제 지리 비율로(경도 1도 < 위도 1도 보정)."""
    return 1.0 / math.cos(math.radians(SEOUL_LAT))

def style_map_ax(ax):
    ax.set_aspect(map_aspect())
    ax.axis("off")
