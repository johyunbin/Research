# -*- coding: utf-8 -*-
"""
Paper32 — Figure 1: PRISMA 2020 흐름도 (2갈래 = 데이터베이스 검색 + 인용추적)
등록 프로토콜의 "other search strategies"를 이행했으므로 PRISMA 2020 표준 2열 양식으로 그린다.
두 갈래 모두 전문심사까지 완료(2026-08-03) — 최종 포함 96편.
수치 출처: fulltext/prisma_flow.md (DB 갈래) · ct_screen_final.csv·ct_retrieval_status.csv (인용추적)
출력: figures/Fig1_PRISMA.png|pdf
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
INK, ACC, MUT, POS = T.INK, T.BLUE, T.AXIS, T.BLUE
GREY = "#f6f6f4"


def main():
    fig, ax = plt.subplots(figsize=(10.2, 8.4))
    ax.set_xlim(0, 12.2); ax.set_ylim(0, 10.6); ax.axis("off")
    fig.subplots_adjust(left=0.008, right=0.992, top=0.992, bottom=0.008)

    # 열 좌표: [DB 본류][DB 배제] | [인용추적 본류][인용추적 배제]
    AX, AW = 0.72, 3.55
    BX, BW = 4.40, 2.58
    CX_, CW = 7.52, 2.55
    DX, DW = 10.15, 1.95
    ACX, CCX = AX + AW / 2, CX_ + CW / 2

    def box(x, y, w, h, text, fc="white", ec=INK, fs=7.6, ls=1.42):
        ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.05,rounding_size=0.08",
                                    fc=fc, ec=ec, lw=0.9))
        ax.text(x + w / 2, y + h / 2, text, ha="center", va="center", fontsize=fs,
                color=INK, linespacing=ls)

    def arr(x1, y1, x2, y2, dashed=False, color=INK):
        ax.add_patch(FancyArrowPatch((x1, y1), (x2, y2), arrowstyle="-|>", mutation_scale=9,
                                     lw=0.9, color=color, shrinkA=0, shrinkB=1,
                                     linestyle=(0, (3, 2)) if dashed else "-"))

    # 단계 밴드
    for y0, y1, lab in [(8.62, 10.42, "Identification"), (5.72, 8.42, "Screening"),
                        (1.62, 5.52, "Eligibility"), (0.12, 1.42, "Included")]:
        ax.add_patch(FancyBboxPatch((0.04, y0), 0.44, y1 - y0, boxstyle="round,pad=0.02",
                                    fc="#eef2f5", ec="none"))
        ax.text(0.26, (y0 + y1) / 2, lab, rotation=90, ha="center", va="center",
                fontsize=8.6, color=ACC, fontweight="bold")

    # 열 제목
    ax.text(ACX, 10.50, "Identification of studies via databases", ha="center",
            fontsize=8.8, fontweight="bold", color=INK)
    ax.text(CCX, 10.50, "Identification of studies via other methods", ha="center",
            fontsize=8.8, fontweight="bold", color=INK)
    ax.plot([7.20, 7.20], [0.10, 10.36], color=T.GRID, lw=0.9, zorder=0)

    # ── A. 데이터베이스 갈래 ────────────────────────────────────────
    box(AX, 9.38, AW, 0.98,
        "Records identified from databases (2 Aug 2026)\n"
        "Web of Science 1,010 · Scopus 850 · PubMed 213\n$\\bf{n\\ =\\ 2{,}073}$")
    box(BX, 9.55, BW, 0.62, "Duplicates removed\nn = 757", fc=GREY, ec=MUT, fs=7.2)
    arr(AX + AW, 9.87, BX, 9.87)
    arr(ACX, 9.38, ACX, 8.42)

    box(AX, 7.62, AW, 0.72, "Records screened (title / abstract)\n$\\bf{n\\ =\\ 1{,}316}$")
    box(BX, 6.62, BW, 1.72,
        "$\\bf{Records\\ excluded\\ (n = 1{,}127)}$\n"
        "Animal / wildlife   479\nPerception or health only   325\n"
        "Setting not eligible   133\nNot empirical   109\nNo acoustic variable   81",
        fc=GREY, ec=MUT, fs=6.8)
    arr(AX + AW, 7.98, BX, 7.98)
    arr(ACX, 7.62, ACX, 5.52)

    box(AX, 4.72, AW, 0.72, "Reports sought for retrieval\n$\\bf{n\\ =\\ 189}$")
    box(BX, 4.02, BW, 1.32,
        "$\\bf{Not\\ retrieved\\ (n = 89)}$\n"
        "No institutional access   84\nPay-per-view only   3\nAbstract only   2",
        fc=GREY, ec=MUT, fs=6.8)
    arr(AX + AW, 5.08, BX, 5.08)
    arr(ACX, 4.72, ACX, 3.92)

    box(AX, 3.20, AW, 0.72, "Reports assessed for eligibility\n$\\bf{n\\ =\\ 100}$")
    box(BX, 2.10, BW, 1.62,
        "$\\bf{Reports\\ excluded\\ (n = 16)}$\n"
        "Setting not eligible   7\nNo observable behaviour   6\n"
        "Perceptual outcome only   1\nNo acoustic exposure   1\nNot empirical   1",
        fc=GREY, ec=MUT, fs=6.8)
    arr(AX + AW, 3.56, BX, 3.56)
    box(BX, 1.40, BW, 0.62, "Reserved for sensitivity analysis\n(intention outcomes)   n = 3",
        fc="#fbf7ec", ec=MUT, fs=6.8)
    arr(AX + AW, 3.30, BX, 1.74)
    arr(ACX, 3.20, ACX, 1.42)

    # ── C. 인용추적 갈래 ────────────────────────────────────────────
    box(CX_, 9.38, CW, 0.98,
        "Records identified by citation searching\n"
        "84 included studies as seeds (OpenAlex)\n"
        "backward 413 · forward 1,660\n$\\bf{n\\ =\\ 2{,}073}$", fs=7.2, ls=1.36)
    box(DX, 9.30, DW, 1.06,
        "$\\bf{Deprioritised\\ (n = 1{,}645)}$\n"
        "Registered 3-block logic\napplied to titles\nAnimal / acoustic ecology   26\n"
        "≤1 block matched   1,619", fc=GREY, ec=MUT, fs=6.2, ls=1.34)
    arr(CX_ + CW, 9.83, DX, 9.83)
    arr(CCX, 9.38, CCX, 8.42)

    box(CX_, 7.62, CW, 0.72, "Records screened (title)\n$\\bf{n\\ =\\ 428}$")
    box(DX, 6.72, DW, 1.62,
        "$\\bf{Excluded\\ (n = 282)}$\n"
        "No behavioural outcome   211\nNo acoustic variable   29\n"
        "Not empirical   26\nSetting not eligible   10\nAnimal   6",
        fc=GREY, ec=MUT, fs=6.2)
    arr(CX_ + CW, 7.98, DX, 7.98)
    arr(CCX, 7.62, CCX, 6.92)

    box(CX_, 6.20, CW, 0.72, "Records screened (abstract)\n$\\bf{n\\ =\\ 146}$")
    box(DX, 5.50, DW, 1.12,
        "$\\bf{Excluded\\ (n = 67)}$\n"
        "No acoustic variable   37\nNo behavioural outcome   26\nOther   4",
        fc=GREY, ec=MUT, fs=6.2)
    arr(CX_ + CW, 6.56, DX, 6.56)
    arr(CCX, 6.20, CCX, 5.52)

    box(CX_, 4.72, CW, 0.72, "Reports sought for retrieval\n$\\bf{n\\ =\\ 79}$")
    box(DX, 4.02, DW, 1.32,
        "$\\bf{Not\\ retrieved\\ (n = 25)}$\n"
        "No institutional access   22\nPay-per-view only   3",
        fc=GREY, ec=MUT, fs=6.2)
    arr(CX_ + CW, 5.08, DX, 5.08)
    arr(CCX, 4.72, CCX, 3.92)

    box(CX_, 3.20, CW, 0.72, "Reports assessed for eligibility\n$\\bf{n\\ =\\ 54}$")
    box(DX, 2.10, DW, 1.62,
        "$\\bf{Reports\\ excluded\\ (n = 38)}$\n"
        "No acoustic variable   20\nNo observable behaviour   17\n"
        "Not in English (Chinese)   1",
        fc=GREY, ec=MUT, fs=6.2)
    arr(CX_ + CW, 3.56, DX, 3.56)
    box(DX, 1.40, DW, 0.62, "Sensitivity analysis only\nn = 1", fc="#fbf7ec", ec=MUT, fs=6.2)
    arr(CX_ + CW, 3.30, DX, 1.74)
    arr(CCX, 3.20, CCX, 1.42)

    # ── Included ───────────────────────────────────────────────────
    ax.add_patch(FancyBboxPatch((AX, 0.18), 9.5, 1.06,
                                boxstyle="round,pad=0.05,rounding_size=0.08",
                                fc="#eaf3ee", ec=POS, lw=1.1))
    ax.text(AX + 9.5 / 2, 0.94, "Studies included in the review    $\\bf{n\\ =\\ 96}$",
            ha="center", va="center", fontsize=9.0, color=INK)
    ax.text(AX + 9.5 / 2, 0.62, "via databases  n = 81          via citation searching  n = 15"
            "          reserved for sensitivity analysis  n = 4",
            ha="center", va="center", fontsize=7.4, color=INK)
    ax.text(AX + 9.5 / 2, 0.36, "Citation searching added 15 studies (19% of the database yield) — "
            "the registered supplementary route was not redundant.",
            ha="center", va="center", fontsize=6.8, color=T.INK2)

    for ext in ("png", "pdf"):
        fig.savefig(os.path.join(FIG, f"Fig1_PRISMA.{ext}"), dpi=300, bbox_inches=None)
    plt.close(fig)
    print("[저장] figures/Fig1_PRISMA.png|pdf (2갈래 PRISMA 2020)")


if __name__ == "__main__":
    main()
