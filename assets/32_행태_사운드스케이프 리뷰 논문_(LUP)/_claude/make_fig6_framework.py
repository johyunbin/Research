# -*- coding: utf-8 -*-
"""
Paper32 — Figure 6: 개념 프레임워크 (등록 osf.io/7ew8q 산출물 약속 이행)
"The soundscape-behaviour loop in public open space"
  · 양방향 루프(forward 음→행태 / reverse 행태→음)
  · 관여 경사(engagement gradient) 5단 — 밴드 농도 = 풀링된 효과크기 수(k), 캡션에 명시
  · 행태 측정 3세대(자기보고 → 체계적 관찰 → 센싱·궤적)
viz_theme 검증 팔레트 · 게재용이므로 도면 텍스트는 전부 영문.
출력: figures/Fig6_Framework.png|pdf
"""
import sys, os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch, Rectangle
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import viz_theme as T

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
FIG = os.path.join(BASE, "figures")
os.makedirs(FIG, exist_ok=True)
T.apply()

# ★ 수치는 전부 figures/fig_data.json 에서 읽는다(구판은 81편·구 MA 값을 하드코딩했다).
import json as _json
_D = _json.load(open(os.path.join(FIG, "fig_data.json"), encoding="utf-8"))


def _ma(key):
    b = _D["ma"][key]; p = b["pooled"]; q = b["quality_mix"]
    SH = {"high": "high", "moderate": "mod", "low": "low"}
    qs = " · ".join(f"{SH[k]} {v}" for k, v in q.items() if v) or "—"
    star = "**" if p["p"] < .01 else ("*" if p["p"] < .05 else "")
    if b.get("back_r"):
        import math as _m
        val = f"r = {_m.tanh(p['est']):+.2f} {star}".strip()
    else:
        val = f"g = {p['est']:+.2f} {star}".strip()
    return f"k = {p['k']}", val, qs, p["k"]


FWD, REV = T.BLUE, T.ORANGE


def box(ax, cx, cy, w, h, fc, ec, lw=1.1, z=3):
    ax.add_patch(FancyBboxPatch((cx - w / 2, cy - h / 2), w, h,
                                boxstyle="round,pad=0,rounding_size=1.8",
                                facecolor=fc, edgecolor=ec, linewidth=lw, zorder=z))


def arrow(ax, p0, p1, color, rad=0.0, lw=2.0, z=4, ms=15):
    ax.add_patch(FancyArrowPatch(p0, p1, connectionstyle=f"arc3,rad={rad}",
                                 arrowstyle="-|>", mutation_scale=ms, linewidth=lw,
                                 color=color, zorder=z, shrinkA=2, shrinkB=2))


def main():
    fig, ax = plt.subplots(figsize=(T.W_FULL, 5.1))
    # 세대 블록 삭제로 y<24 가 비었다 — 그만큼 잘라 여백을 없앤다
    ax.set_xlim(0, 100); ax.set_ylim(23, 101); ax.axis("off")

    # ── 상단: 양방향 루프 ────────────────────────────────────────────
    box(ax, 50, 93, 40, 11, T.TINT_BLUE, FWD, 1.2)
    ax.text(50, 95.6, "ACOUSTIC ENVIRONMENT", ha="center", va="center",
            fontsize=8.8, fontweight="bold", color=T.INK, zorder=5)
    ax.text(50, 91.0, "source composition · level · temporal variation\n"
            "traffic  ·  natural  ·  music  ·  human",
            ha="center", va="center", fontsize=6.8, color=T.INK2, zorder=5, linespacing=1.5)

    box(ax, 84, 74, 28, 11, T.SURF, T.BASE, 1.0)
    ax.text(84, 76.6, "APPRAISAL", ha="center", va="center",
            fontsize=8.8, fontweight="bold", color=T.INK, zorder=5)
    ax.text(84, 72.0, "pleasantness · eventfulness\n(ISO 12913) · expectancy fit",
            ha="center", va="center", fontsize=6.8, color=T.INK2, zorder=5, linespacing=1.5)

    box(ax, 16, 74, 28, 11, T.SURF, T.BASE, 1.0)
    ax.text(16, 76.6, "ACTIVITY & OCCUPANCY", ha="center", va="center",
            fontsize=8.0, fontweight="bold", color=T.INK, zorder=5)
    ax.text(16, 72.0, "who · how many · doing what\n(people generate sound)",
            ha="center", va="center", fontsize=6.8, color=T.INK2, zorder=5, linespacing=1.5)

    box(ax, 50, 55, 36, 10.5, T.TINT_TERRA, REV, 1.2)
    ax.text(50, 57.5, "BEHAVIOURAL RESPONSE", ha="center", va="center",
            fontsize=8.8, fontweight="bold", color=T.INK, zorder=5)
    ax.text(50, 53.2, "observable behaviour — unfolded below",
            ha="center", va="center", fontsize=6.8, color=T.INK2, zorder=5)

    box(ax, 50, 74, 30, 13.5, T.TINT_NEUT, T.BASE, 0.9, z=2)
    ax.text(50, 78.6, "MODERATORS", ha="center", va="center", fontsize=7.8,
            fontweight="bold", color=T.INK2, zorder=5)
    ax.text(50, 73.0, "setting type (park · street · square)\nvisual–acoustic congruence\n"
            "purpose of stay · culture · person",
            ha="center", va="center", fontsize=6.6, color=T.INK2, zorder=5, linespacing=1.7)

    arrow(ax, (68.5, 89.2), (79.5, 80.4), FWD, rad=-0.22, lw=2.0)
    arrow(ax, (83.5, 68.2), (68.5, 58.4), FWD, rad=-0.24, lw=2.0)
    arrow(ax, (31.5, 58.4), (16.5, 68.2), REV, rad=-0.24, lw=2.0)
    arrow(ax, (20.5, 80.4), (31.5, 89.2), REV, rad=-0.22, lw=2.0)

    ax.text(79.8, 86.4, "FORWARD", fontsize=7.8, fontweight="bold", color=FWD,
            ha="center", rotation=-40, zorder=6)
    ax.text(20.2, 86.4, "REVERSE", fontsize=7.8, fontweight="bold", color=REV,
            ha="center", rotation=40, zorder=6)

    # ── 중단: 관여 경사 ─────────────────────────────────────────────
    GY, GX0, GX1 = 34.5, 7, 93
    # shade = 풀링된 효과크기 수 k / 5 (캡션에 명시 — 라벨 없는 색 인코딩 금지)
    _w = _ma("walking"); _s = _ma("staying"); _c = _ma("social"); _r = _ma("correlation")
    steps = [
        ("Avoid",       "speed up · leave",   f"MA1   {_w[0]}", _w[1], _w[2], _w[3]),
        ("Pass",        "walk through",       "no pooling",     "narrative only", "—", 0),
        ("Linger",      "stay · sit",         f"MA2   {_s[0]}", _s[1], _s[2], _s[3]),
        ("Interact",    "talk · group",       f"MA3   {_c[0]}", _c[1], _c[2], _c[3]),
        ("Appropriate", "occupy · use space", f"MA4   {_r[0]}", _r[1], _r[2], _r[3]),
    ]
    n = len(steps); gap = 1.6
    wstep = (GX1 - GX0 - gap * (n - 1)) / n

    ax.text(GX0, 46.6, "Engagement gradient", fontsize=9.0, fontweight="bold",
            color=T.INK, ha="left")
    ax.text(GX1, 46.6, "less engaged  →  more engaged", fontsize=6.8,
            color=T.INK2, ha="right")
    arrow(ax, (50, 49.4), (50, 42.2), T.BASE, lw=1.2, z=2, ms=11)

    for i, (name, sub, kline, eline, q, k) in enumerate(steps):
        x0 = GX0 + i * (wstep + gap)
        col = T.SEQ(0.22 + 0.68 * min(k / 6, 1.0))
        fg = T.ink_on(col)                       # 배경 휘도로 글자색 결정(눈대중 금지)
        fg2 = "#DCE7F1" if fg == T.SURF else T.INK2
        ax.add_patch(Rectangle((x0, GY - 6.5), wstep, 13.0, facecolor=col,
                               edgecolor=T.BASE if k < 3 else "none",
                               linewidth=0.8, zorder=3))
        ax.text(x0 + wstep / 2, GY + 4.5, name, ha="center", va="center", fontsize=8.5,
                fontweight="bold", color=fg, zorder=5)
        ax.text(x0 + wstep / 2, GY + 1.6, sub, ha="center", va="center", fontsize=6.4,
                color=fg2, zorder=5)
        ax.text(x0 + wstep / 2, GY - 1.8, kline, ha="center", va="center", fontsize=7.2,
                fontweight="bold", color=fg, zorder=5)
        ax.text(x0 + wstep / 2, GY - 4.4, eline, ha="center", va="center", fontsize=7.2,
                fontweight="bold", color=fg, zorder=5)
        ax.text(x0 + wstep / 2, GY - 9.0, q, ha="center", va="center", fontsize=6.6,
                color=T.INK2, zorder=5)


    # ★ '측정 3세대' 블록은 삭제했다 — 같은 내용을 Fig5_Methods 가 이미 담고 있어
    #   한 그림에 세 덩어리를 넣은 것이 이 도면을 읽기 어렵게 만든 주된 원인이었다.

    # 제목·부제는 그림에 넣지 않는다(캡션이 담당).

    fig.subplots_adjust(left=0.02, right=0.98, top=0.985, bottom=0.015)
    for ext in ("png", "pdf"):
        fig.savefig(os.path.join(FIG, f"Fig6_Framework.{ext}"), dpi=300)
    plt.close(fig)
    print("[저장] figures/Fig6_Framework.png|pdf")


if __name__ == "__main__":
    main()
