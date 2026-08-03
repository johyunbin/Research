# -*- coding: utf-8 -*-
"""
Paper32 — 그림 공통 테마 (검증된 팔레트 · 저널 투고 규격)
검증(OKLab ΔE + deuteranopia 시뮬, validate_palette 기준 재현):
  blue #2a78d6 vs orange #eb6834 : normal 33.6 / CVD 31.8   PASS
  blue vs muted  #898781         : normal 17.8 / CVD 18.3   PASS
  orange vs muted                : normal 17.6 / CVD 13.6   PASS
  (구 팔레트 #2c6e91 vs #8a8a8a = normal 14.9 → 15 floor 미달로 폐기)
순차(sequential)는 단일 hue blue 램프 100→700 사용(다색 램프 YlGnBu 폐기 — 무지개 금지 규칙).
"""
import matplotlib as mpl
from matplotlib.colors import LinearSegmentedColormap

# 카테고리(고정 순서 — 순환 금지)
SERIES = ["#2a78d6", "#eb6834", "#898781"]
BLUE, ORANGE, MUTED = SERIES

# 잉크·크롬
INK = "#0b0b0b"      # primary
INK2 = "#52514e"     # secondary
AXIS = "#898781"     # muted (축·라벨)
GRID = "#e1e0d9"     # hairline
BASE = "#c3c2b7"     # baseline
SURF = "#ffffff"     # 인쇄 대비 백색 표면

# 단일 hue 순차 램프 (blue 100→700)
SEQ_STEPS = ["#cde2fb", "#b7d3f6", "#9ec5f4", "#86b6ef", "#6da7ec", "#5598e7",
             "#3987e5", "#2a78d6", "#256abf", "#1c5cab", "#184f95", "#104281", "#0d366b"]
SEQ = LinearSegmentedColormap.from_list("p32_blue", SEQ_STEPS)

FRAME_POS = "#1c5cab"   # 강조(딥 블루)
FRAME_NEG = "#eb6834"


def apply():
    mpl.rcParams.update({
        "font.family": "DejaVu Sans",
        "font.size": 9,
        "text.color": INK,
        "axes.edgecolor": BASE,
        "axes.labelcolor": INK2,
        "axes.linewidth": 0.8,
        "axes.titlecolor": INK,
        "xtick.color": AXIS, "ytick.color": AXIS,
        "xtick.labelcolor": INK2, "ytick.labelcolor": INK2,
        "xtick.major.width": 0.7, "ytick.major.width": 0.7,
        "grid.color": GRID, "grid.linewidth": 0.7,
        "figure.facecolor": SURF, "axes.facecolor": SURF,
        "savefig.facecolor": SURF,
        "savefig.dpi": 300, "savefig.bbox": "tight",
        "legend.frameon": False,
    })
