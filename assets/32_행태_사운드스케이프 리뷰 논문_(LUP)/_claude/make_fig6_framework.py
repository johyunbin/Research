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
    fig, ax = plt.subplots(figsize=(8.6, 9.8))
    ax.set_xlim(0, 100); ax.set_ylim(0, 100); ax.axis("off")

    # ── 상단: 양방향 루프 ────────────────────────────────────────────
    box(ax, 50, 93, 32, 11, "#eef4fd", FWD, 1.3)
    ax.text(50, 95.6, "ACOUSTIC ENVIRONMENT", ha="center", va="center",
            fontsize=9.6, fontweight="bold", color=T.INK, zorder=5)
    ax.text(50, 91.0, "source composition · level · temporal variation\n"
            "traffic  ·  natural  ·  music  ·  human",
            ha="center", va="center", fontsize=7.2, color=T.INK2, zorder=5, linespacing=1.5)

    box(ax, 84, 74, 27, 11, T.SURF, T.BASE, 1.0)
    ax.text(84, 76.6, "APPRAISAL", ha="center", va="center",
            fontsize=9.6, fontweight="bold", color=T.INK, zorder=5)
    ax.text(84, 72.0, "pleasantness · eventfulness\n(ISO 12913) · expectancy fit",
            ha="center", va="center", fontsize=7.2, color=T.INK2, zorder=5, linespacing=1.5)

    box(ax, 16, 74, 27, 11, T.SURF, T.BASE, 1.0)
    ax.text(16, 76.6, "ACTIVITY & OCCUPANCY", ha="center", va="center",
            fontsize=8.5, fontweight="bold", color=T.INK, zorder=5)
    ax.text(16, 72.0, "who · how many · doing what\n(people generate sound)",
            ha="center", va="center", fontsize=7.2, color=T.INK2, zorder=5, linespacing=1.5)

    box(ax, 50, 55, 36, 10.5, "#fdefe8", REV, 1.3)
    ax.text(50, 57.5, "BEHAVIOURAL RESPONSE", ha="center", va="center",
            fontsize=9.6, fontweight="bold", color=T.INK, zorder=5)
    ax.text(50, 53.2, "observable behaviour — unfolded below",
            ha="center", va="center", fontsize=7.2, color=T.INK2, zorder=5)

    box(ax, 50, 74, 28, 13.5, "#faf9f4", T.BASE, 0.9, z=2)
    ax.text(50, 78.6, "MODERATORS", ha="center", va="center", fontsize=8.3,
            fontweight="bold", color=T.INK2, zorder=5)
    ax.text(50, 73.0, "setting type (park · street · square)\nvisual–acoustic congruence\n"
            "purpose of stay · culture · person",
            ha="center", va="center", fontsize=7.0, color=T.INK2, zorder=5, linespacing=1.7)

    arrow(ax, (66.5, 90.4), (78.5, 80.2), FWD, rad=-0.24, lw=2.2)
    arrow(ax, (83.5, 68.2), (68.5, 58.4), FWD, rad=-0.24, lw=2.2)
    arrow(ax, (31.5, 58.4), (16.5, 68.2), REV, rad=-0.24, lw=2.2)
    arrow(ax, (21.5, 80.2), (33.5, 90.4), REV, rad=-0.24, lw=2.2)

    ax.text(78.2, 86.2, "FORWARD", fontsize=8.2, fontweight="bold", color=FWD,
            ha="center", rotation=-38, zorder=6)
    ax.text(21.8, 86.2, "REVERSE", fontsize=8.2, fontweight="bold", color=REV,
            ha="center", rotation=38, zorder=6)

    # ── 중단: 관여 경사 ─────────────────────────────────────────────
    GY, GX0, GX1 = 34.5, 7, 93
    # shade = 풀링된 효과크기 수 k / 5 (캡션에 명시 — 라벨 없는 색 인코딩 금지)
    steps = [
        ("Avoid",       "speed up · leave",   "MA1   k = 4", "g = −0.50",     "low 4 / 4",      4),
        ("Pass",        "walk through",       "no pooling",  "narrative only", "—",             0),
        ("Linger",      "stay · sit",         "MA2   k = 3", "g = +0.31",     "high 1 · mod 2", 3),
        ("Interact",    "talk · group",       "MA3   k = 4", "g = +0.65 *",   "high 2 · mod 2", 4),
        ("Appropriate", "occupy · use space", "MA4   k = 7", "r = +0.41 **",  "high 2 · mod 4", 7),
    ]
    n = len(steps); gap = 1.6
    wstep = (GX1 - GX0 - gap * (n - 1)) / n

    ax.text(GX0, 46.6, "Engagement gradient", fontsize=9.4, fontweight="bold",
            color=T.INK, ha="left")
    ax.text(GX1, 46.6, "less engaged   ————→   more engaged", fontsize=7.2,
            color=T.INK2, ha="right")
    arrow(ax, (50, 49.4), (50, 42.2), T.BASE, lw=1.2, z=2, ms=11)

    for i, (name, sub, kline, eline, q, k) in enumerate(steps):
        x0 = GX0 + i * (wstep + gap)
        col = T.SEQ(0.14 + 0.58 * min(k / 7, 1.0))
        fg = T.SURF if k >= 3 else T.INK
        fg2 = "#e9f1fc" if k >= 3 else T.INK2
        ax.add_patch(Rectangle((x0, GY - 6.5), wstep, 13.0, facecolor=col,
                               edgecolor=T.BASE if k < 3 else "none",
                               linewidth=0.8, zorder=3))
        ax.text(x0 + wstep / 2, GY + 4.5, name, ha="center", va="center", fontsize=9.0,
                fontweight="bold", color=fg, zorder=5)
        ax.text(x0 + wstep / 2, GY + 1.6, sub, ha="center", va="center", fontsize=6.6,
                color=fg2, zorder=5)
        ax.text(x0 + wstep / 2, GY - 1.8, kline, ha="center", va="center", fontsize=7.4,
                fontweight="bold", color=fg, zorder=5)
        ax.text(x0 + wstep / 2, GY - 4.4, eline, ha="center", va="center", fontsize=7.4,
                fontweight="bold", color=fg, zorder=5)
        ax.text(x0 + wstep / 2, GY - 9.0, q, ha="center", va="center", fontsize=6.9,
                color=T.INK2, zorder=5)

    ax.text(GX0, GY - 12.4, "Band shading = number of pooled effect sizes (k); * p < .05, ** p < .01. "
            "Text below each band = MMAT quality of the contributing studies.",
            fontsize=6.9, color=T.INK2, ha="left")
    ax.text(GX0, GY - 15.0, "All four walking-speed effects come from low-quality studies — "
            "excluding them leaves nothing to pool (k = 0).",
            fontsize=6.9, color=REV, ha="left", fontweight="bold")

    # ── 하단: 측정 3세대 ────────────────────────────────────────────
    MY = 6.2
    ax.plot([GX0, GX1], [16.4, 16.4], color=T.GRID, lw=0.9, zorder=1)
    ax.text(GX0, 14.2, "Generations of behavioural measurement", fontsize=9.4,
            fontweight="bold", color=T.INK, ha="left")
    ax.text(GX0, 12.0, "Generations accumulate rather than replace one another: "
            "G3 rose from 3 studies (2010s) to 16 (2020s) while G1 and G2 also grew. "
            "Studies using two generations are counted in both.",
            fontsize=6.9, color=T.INK2, ha="left")

    # 카운트는 Table 1의 measure_gen에서 산출(하드코딩 금지 — Fig 5와 같은 규칙).
    from collections import Counter
    gc = Counter()
    with open(os.path.join(FT, "table1_v2.csv"), encoding="utf-8-sig") as f:
        import csv as _csv
        for r in _csv.DictReader(f):
            for g in (x.strip() for x in r["measure_gen"].split(";")):
                if g in ("G1", "G2", "G3"):
                    gc[g] += 1
    gens = [("G1", "Self-report", "survey · recall", str(gc["G1"]), 0.30),
            ("G2", "Systematic\nobservation", "mapping · counts", str(gc["G2"]), 0.58),
            ("G3", "Sensing &\ntrajectory", "GPS · video · ML", str(gc["G3"]), 0.88)]
    gw = (GX1 - GX0 - 2 * 2.4) / 3
    CHIP = 10.0
    for i, (tag, name, sub, cnt, shade) in enumerate(gens):
        x0 = GX0 + i * (gw + 2.4)
        c = T.SEQ(0.2 + 0.6 * shade)
        ax.add_patch(Rectangle((x0, MY - 4.4), gw, 9.0, facecolor=T.SURF,
                               edgecolor=T.BASE, linewidth=0.9, zorder=3))
        ax.add_patch(Rectangle((x0, MY - 4.4), CHIP, 9.0, facecolor=c, edgecolor="none", zorder=4))
        ax.text(x0 + CHIP / 2, MY + 2.7, tag, fontsize=7.8, fontweight="bold",
                color=T.SURF, ha="center", va="center", zorder=5)
        ax.text(x0 + CHIP / 2, MY - 0.6, cnt, fontsize=12.0, fontweight="bold",
                color=T.SURF, ha="center", va="center", zorder=5)
        ax.text(x0 + CHIP / 2, MY - 3.3, "studies", fontsize=6.0,
                color="#e9f1fc", ha="center", va="center", zorder=5)
        ax.text(x0 + CHIP + 2.4, MY + 1.8, name, fontsize=7.8, fontweight="bold",
                color=T.INK, ha="left", va="center", zorder=5, linespacing=1.35)
        ax.text(x0 + CHIP + 2.4, MY - 2.6, sub, fontsize=6.4, color=T.INK2, ha="left",
                va="center", zorder=5)

    fig.text(0.045, 0.982, "The soundscape–behaviour loop in public open space",
             fontsize=12.6, fontweight="bold", color=T.INK, ha="left", va="top")
    fig.text(0.045, 0.960, "A conceptual framework linking the acoustic environment, behaviour, "
             "and generations of measurement,\nwith the 81-study evidence base mapped onto it. "
             "Blue paths: sound shapes behaviour. Orange paths: activity shapes sound.",
             fontsize=8.0, color=T.INK2, ha="left", va="top", linespacing=1.5)

    fig.subplots_adjust(left=0.02, right=0.98, top=0.925, bottom=0.012)
    for ext in ("png", "pdf"):
        fig.savefig(os.path.join(FIG, f"Fig6_Framework.{ext}"), dpi=300)
    plt.close(fig)
    print("[저장] figures/Fig6_Framework.png|pdf")


if __name__ == "__main__":
    main()
