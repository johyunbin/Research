# -*- coding: utf-8 -*-
"""
Paper32 — Figure 8(본문 번호): 상호적 증거 프레임워크 (reciprocal evidence framework)

v3 (2026-09-16 외부 AI 검토 반영)
  · 닫힌 인과 피드백 고리의 "증명"이 아니라 **상호적 증거 구조**를 그리는 개념도로 재정의
  · 공간적·물리적·사회문화적 **맥락 층**을 틀 전체를 감싸는 프레임으로 추가 — 맥락이 두 경로를
    모두 조절한다는 뜻을 화살표 없이 포함 관계로 표현(검토본은 떠 있는 주석이었다)
  · 음환경 → 행태의 **직접 경로**(평가를 거치지 않는)를 점선으로 추가
  · 역방향을 활동·점유 → 소리 발생 → 음환경의 두 단계로 분리
  · ★ **정량 정보(k·g·p·MMAT 음영) 전부 제거** — 개념도는 개념만, 수치는 Results·표·forest 담당
  · 관여의 경사는 **종합 장치(synthesis device)이지 검증된 서열 척도가 아님**을 도면에 명시.
    서열 척도로 오독되지 않게 칸에 단계형 색 농도를 쓰지 않는다
검토본 도면(Manuscript_KO_20260828_FINAL 내 이미지)은 문구가 SOUND PRODUCTION 상자를
덮고 되돌아가는 화살표가 상자를 관통해 채택하지 않고, 개념만 받아 여기서 다시 그렸다.
출력: figures/Fig6_Framework.png|pdf (파일명은 빌더 호환을 위해 유지)
"""
import sys, os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import viz_theme as T

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FIG = os.path.join(BASE, "figures")
os.makedirs(FIG, exist_ok=True)
T.apply()

FWD, REV = T.BLUE, T.TERRA


def box(ax, cx, cy, w, h, title, sub, fc=T.SURF, ec=T.BASE, lw=1.0, fs=8.0, z=3):
    ax.add_patch(FancyBboxPatch((cx - w / 2, cy - h / 2), w, h,
                                boxstyle="round,pad=0,rounding_size=1.6",
                                facecolor=fc, edgecolor=ec, linewidth=lw, zorder=z))
    ax.text(cx, cy + h * 0.17, title, ha="center", va="center", fontsize=fs,
            fontweight="bold", color=T.INK, zorder=z + 1)
    ax.text(cx, cy - h * 0.22, sub, ha="center", va="center", fontsize=6.3,
            color=T.INK2, zorder=z + 1)


def arrow(ax, p0, p1, color, rad=0.0, lw=1.8, ls="-", ms=12, z=4):
    ax.add_patch(FancyArrowPatch(p0, p1, connectionstyle=f"arc3,rad={rad}",
                                 arrowstyle="-|>", mutation_scale=ms, linewidth=lw,
                                 linestyle=ls, color=color, zorder=z,
                                 shrinkA=1.5, shrinkB=1.5))


def main():
    fig, ax = plt.subplots(figsize=(T.W_FULL, 5.0))
    fig.subplots_adjust(left=0, right=1, top=1, bottom=0)
    ax.set_xlim(0, 100); ax.set_ylim(0, 100); ax.axis("off")

    # ── 맥락 프레임: 틀 전체를 감싼다 = 두 경로 모두를 조절 ─────────────
    ax.add_patch(FancyBboxPatch((1.2, 23.0), 97.6, 75.8,
                                boxstyle="round,pad=0,rounding_size=2.2",
                                facecolor=T.TINT_NEUT, edgecolor=T.BASE, linewidth=0.9, zorder=1))
    ax.text(3.6, 95.2, "SPATIAL · PHYSICAL · SOCIO-CULTURAL CONTEXT", ha="left", va="center",
            fontsize=7.6, fontweight="bold", color=T.INK2, zorder=2)
    ax.text(3.6, 91.4, "setting type · visual–acoustic congruence · purpose of stay · "
            "time · culture", ha="left", va="center", fontsize=6.2, color=T.INK2, zorder=2)
    ax.text(96.4, 95.2, "moderates both pathways", ha="right", va="center", fontsize=6.4,
            style="italic", color=T.INK2, zorder=2)

    # ── 노드 ─────────────────────────────────────────────────────────
    AE = (50, 79.5); APP = (81.5, 57.5); BEH = (50, 34.5); ACT = (18.5, 46.5); SND = (18.5, 68.5)
    box(ax, *AE, 31, 9.5, "ACOUSTIC ENVIRONMENT", "source composition · level · temporal pattern",
        fc=T.TINT_BLUE, ec=FWD, lw=1.2, fs=8.2)
    box(ax, *APP, 28, 9.5, "SOUNDSCAPE APPRAISAL", "perception · interpretation", fs=7.8)
    box(ax, *BEH, 39, 9.5, "OBSERVABLE BEHAVIOUR",
        "movement / passing · staying / space use · interaction",
        fc=T.TINT_TERRA, ec=REV, lw=1.2, fs=8.2)
    box(ax, *ACT, 28, 9.5, "ACTIVITY & OCCUPANCY", "density · programming · companionship",
        fs=7.6)
    box(ax, *SND, 28, 9.5, "SOUND PRODUCTION", "voices · activity sound · amplified sound",
        fs=7.6)

    # ── 순방향(파랑): 음환경 → 평가 → 행태, + 직접 경로(점선) ───────────
    arrow(ax, (AE[0] + 15.5, AE[1] - 1.5), (APP[0], APP[1] + 4.75), FWD, rad=-0.28)
    arrow(ax, (APP[0], APP[1] - 4.75), (BEH[0] + 19.5, BEH[1] + 1.0), FWD, rad=-0.28)
    arrow(ax, (AE[0], AE[1] - 4.75), (BEH[0], BEH[1] + 4.75), FWD, lw=1.1, ls=(0, (4, 3)),
          ms=10)
    ax.text(51.8, 57.0, "direct", ha="left", va="center", fontsize=6.2, style="italic",
            color=FWD, zorder=5)
    ax.text(78.0, 74.5, "FORWARD", ha="center", va="center", fontsize=7.4, fontweight="bold",
            color=FWD, rotation=-36, zorder=5)

    # ── 역방향(테라코타): 행태 → 활동·점유 → 소리 발생 → 음환경 ─────────
    arrow(ax, (BEH[0] - 19.5, BEH[1] + 1.0), (ACT[0], ACT[1] - 4.75), REV, rad=-0.28)
    arrow(ax, (ACT[0], ACT[1] + 4.75), (SND[0], SND[1] - 4.75), REV)
    arrow(ax, (SND[0], SND[1] + 4.75), (AE[0] - 15.5, AE[1] - 1.5), REV, rad=-0.28)
    ax.text(21.5, 83.5, "REVERSE", ha="center", va="center", fontsize=7.4, fontweight="bold",
            color=REV, rotation=36, zorder=5)

    # ── 관여의 경사: 종합 장치 (정량 정보 없음) ───────────────────────────
    arrow(ax, (BEH[0], BEH[1] - 4.75), (BEH[0], 19.6), T.BASE, lw=1.0, ms=9, z=2)
    ax.text(3.0, 17.4, "Engagement gradient", ha="left", va="center", fontsize=8.0,
            fontweight="bold", color=T.INK)
    ax.text(97.0, 17.4, "synthesis device · not a validated ordinal behavioural scale",
            ha="right", va="center", fontsize=6.2, style="italic", color=T.INK2)
    steps = [("Avoidance", "speed up · leave"), ("Passing", "walk through"),
             ("Staying", "linger · sit"), ("Interacting", "talk · help"),
             ("Appropriating", "occupy · adapt")]
    x0, x1, gap = 3.0, 97.0, 1.6
    w = (x1 - x0 - gap * (len(steps) - 1)) / len(steps)
    for i, (name, sub) in enumerate(steps):
        cx = x0 + i * (w + gap) + w / 2
        ax.add_patch(FancyBboxPatch((cx - w / 2, 5.2), w, 8.6,
                                    boxstyle="round,pad=0,rounding_size=1.0",
                                    facecolor=T.SURF, edgecolor=T.BASE, linewidth=0.9, zorder=3))
        ax.text(cx, 10.9, name, ha="center", va="center", fontsize=7.6, fontweight="bold",
                color=T.INK, zorder=4)
        ax.text(cx, 7.4, sub, ha="center", va="center", fontsize=6.0, color=T.INK2, zorder=4)
    ax.text(50, 2.2, "less engaged  →  more engaged", ha="center", va="center", fontsize=6.2,
            color=T.INK2)

    for ext in ("png", "pdf"):
        fig.savefig(os.path.join(FIG, f"Fig6_Framework.{ext}"), dpi=300)
    plt.close(fig)
    print("[저장] figures/Fig6_Framework.png|pdf (reciprocal evidence framework v3)")


if __name__ == "__main__":
    main()
