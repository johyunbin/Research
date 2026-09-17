# -*- coding: utf-8 -*-
"""
Paper32 — Figure 1: PRISMA 2020 흐름도

★ 양식 = PRISMA 2020 flow diagram for new systematic reviews which included searches of databases,
  registers and other sources (Page et al., 2021 템플릿). 사용자 지적(2026-09-17): 미확보 사유
  ("No institution access" 등)나 단계별 부가 설명은 일반 리뷰의 흐름도에 쓰지 않는다 → 템플릿 문구와
  칸 구성을 그대로 따르고, 전문 배제 사유만 적는다.
  - 왼쪽 = 데이터베이스(WoS·Scopus·PubMed) / 오른쪽 = 기타 방법(인용 추적 + OpenAlex 보조 검색)
  - 제목 기준 자동 우선순위 필터로 걸러진 레코드 = 템플릿의 "Records marked as ineligible by automation tools"
  - 민감도 분석 전용 연구는 포함 상자에 따로 적는다(배제가 아니다)
★ 수치는 `figures/fig_data.json` 에서만 읽고, 칸 사이 산술이 맞지 않으면 그리지 않고 멈춘다.
★ 글자 크기는 모든 상자에 공통 — 상자마다 따로 맞추면 크기가 들쭉날쭉해진다.
출력: figures/Fig1_PRISMA.png|pdf
"""
import sys, os, json
from collections import Counter
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Rectangle, FancyArrowPatch

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FIG = os.path.join(BASE, "figures")
sys.path.insert(0, BASE)
import viz_theme as T
from figtext import fit_lines
T.apply()
INK = T.INK

D = json.load(open(os.path.join(FIG, "fig_data.json"), encoding="utf-8"))
DB, CT, SP = D["prisma"]["db"], D["prisma"]["ct"], D["prisma"]["supp"]

# ── 기타 방법 열 = 인용 추적 + 보조 검색 합산 ─────────────────────────────
O = {
    "cit": CT["identified"], "supp": SP["identified"],
    "dup": SP["duplicates"], "auto": CT["deprioritised"],
    "screened": CT["screened_title"] + SP["screened"],
    "excluded": CT["excl_title"] + CT["excl_abs"] + SP["excluded"],
    "sought": CT["sought"] + SP["sought"],
    "not_retrieved": CT["not_retrieved"] + SP["not_retrieved"] + SP["prescreen"],
    "assessed": CT["assessed"] + SP["assessed"],
    "sens": CT["sens"], "included": CT["included"] + SP["included"],
}
_ftx = Counter()
for k, v in CT["ftx"] + (SP.get("ftx") or []):
    _ftx[{"No behavioural outcome": "No observed behaviour"}.get(k, k)] += v
O["ftx"] = _ftx.most_common()
O["ft_excluded"] = sum(v for _, v in O["ftx"])

# ── 칸 사이 산술 검산 ────────────────────────────────────────────────────
for name, c in (("databases", DB), ("other methods", O)):
    assert c["screened"] - c["excluded"] == c["sought"], name
    assert c["sought"] - c["not_retrieved"] == c["assessed"], name
    assert c["assessed"] - c["ft_excluded"] - c["sens"] == c["included"], name
assert sum(v for _, v in DB["ftx"]) == DB["ft_excluded"]
assert DB["identified"] - DB["duplicates"] == DB["screened"]
assert CT["identified"] - CT["deprioritised"] == CT["screened_title"]
assert SP["identified"] - SP["duplicates"] == SP["screened"]
assert DB["included"] + O["included"] == D["n_included"]
assert DB["sens"] + O["sens"] == D["n_sens"]

W, H = T.W_FULL, 6.6
FS = 7.2                                                                  # 글자 크기 상한
LM, LS, RM, RS = (5.0, 22.0), (28.6, 22.0), (52.8, 22.0), (76.4, 23.1)   # (x, 폭)
ROW = {"id": (74.0, 15.5), "scr": (59.0, 8.0), "sought": (44.5, 8.0),
       "assess": (24.5, 13.0), "inc": (4.0, 9.0)}                         # (아래 y, 높이)


def main():
    fig, ax = plt.subplots(figsize=(W, H))
    fig.subplots_adjust(left=0, right=1, top=1, bottom=0)
    ax.set_xlim(0, 100); ax.set_ylim(0, 100); ax.axis("off")

    def rect(x, y, w, h, fc="white", ec=INK, lw=0.8):
        ax.add_patch(Rectangle((x, y), w, h, fc=fc, ec=ec, lw=lw))

    def arrow(p0, p1):
        ax.add_patch(FancyArrowPatch(p0, p1, arrowstyle="-|>", mutation_scale=7,
                                     lw=0.8, color=INK, shrinkA=0, shrinkB=0))

    # (정렬, 열, 행, 줄) — 먼저 모두 모은 뒤 공통 글자 크기로 그린다
    boxes = [
        # 데이터베이스
        ("l", LM, "id", ["Records identified from",
                         f"databases (n = {DB['identified']:,}):",
                         f"  Web of Science (n = {DB['wos']:,})",
                         f"  Scopus (n = {DB['scopus']:,})",
                         f"  PubMed (n = {DB['pubmed']:,})"]),
        ("l", LS, "id", ["Records removed before",
                         "screening:",
                         "  Duplicate records",
                         f"  removed (n = {DB['duplicates']:,})"]),
        ("c", LM, "scr", ["Records screened", f"(n = {DB['screened']:,})"]),
        ("c", LS, "scr", ["Records excluded", f"(n = {DB['excluded']:,})"]),
        ("c", LM, "sought", ["Reports sought", "for retrieval", f"(n = {DB['sought']:,})"]),
        ("c", LS, "sought", ["Reports not retrieved", f"(n = {DB['not_retrieved']:,})"]),
        ("c", LM, "assess", ["Reports assessed", "for eligibility", f"(n = {DB['assessed']:,})"]),
        ("l", LS, "assess", ["Reports excluded:"] + [f"{k} (n = {v})" for k, v in DB["ftx"]]),
        # 기타 방법
        ("l", RM, "id", ["Records identified from:",
                         "  Citation searching",
                         f"  (n = {O['cit']:,})",
                         "  Supplementary search",
                         f"  of OpenAlex (n = {O['supp']:,})"]),
        ("l", RS, "id", ["Records removed before",
                         "screening:",
                         "  Duplicate records",
                         f"  removed (n = {O['dup']:,})",
                         "  Records marked as ineligible",
                         f"  by automation tools (n = {O['auto']:,})"]),
        ("c", RM, "scr", ["Records screened", f"(n = {O['screened']:,})"]),
        ("c", RS, "scr", ["Records excluded", f"(n = {O['excluded']:,})"]),
        ("c", RM, "sought", ["Reports sought", "for retrieval", f"(n = {O['sought']:,})"]),
        ("c", RS, "sought", ["Reports not retrieved", f"(n = {O['not_retrieved']:,})"]),
        ("c", RM, "assess", ["Reports assessed", "for eligibility", f"(n = {O['assessed']:,})"]),
        ("l", RS, "assess", ["Reports excluded:"] + [f"{k} (n = {v})" for k, v in O["ftx"]]),
    ]
    (ix, _), (rx, rw) = LM, RM
    iy, ih = ROW["inc"]
    inc_lines = [f"Studies included in review (n = {D['n_included']})",
                 f"Studies used only in sensitivity analyses (n = {D['n_sens']})"]

    f = min([fit_lines(ax, ls, col[1] - 2.2, FS) for _, col, _, ls in boxes]
            + [fit_lines(ax, inc_lines, rx + rw - ix - 2.2, FS)])

    # 머리 띠
    for (x, _), (x2, w2), label in ((LM, LS, "Identification of studies via databases"),
                                    (RM, RS, "Identification of studies via other methods")):
        rect(x, 92.5, x2 + w2 - x, 5.5, fc=T.TINT_NEUT, ec=INK)
        ax.text((x + x2 + w2) / 2, 95.25, label, ha="center", va="center",
                fontsize=f + 0.6, fontweight="bold", color=INK)
    # 단계 띠
    for y0, y1, lab in ((72.5, 90.5, "Identification"), (22.5, 70.5, "Screening"),
                        (4.0, 13.0, "Included")):
        rect(0.4, y0, 3.4, y1 - y0, fc=T.TINT_NEUT, ec=INK, lw=0.6)
        ax.text(2.1, (y0 + y1) / 2, lab, rotation=90, ha="center", va="center",
                fontsize=f, fontweight="bold", color=INK)

    for kind, (x, w), row, ls in boxes:
        y, h = ROW[row]
        rect(x, y, w, h)
        if kind == "c":
            ax.text(x + w / 2, y + h / 2, "\n".join(ls), ha="center", va="center",
                    fontsize=f, color=INK, linespacing=1.45)
        else:
            ax.text(x + 1.1, y + h / 2, "\n".join(ls), ha="left", va="center",
                    fontsize=f, color=INK, linespacing=1.45)
    rect(ix, iy, rx + rw - ix, ih)
    ax.text((ix + rx + rw) / 2, iy + ih / 2, "\n".join(inc_lines), ha="center", va="center",
            fontsize=f, color=INK, linespacing=1.6)

    # 화살표
    for (x, w), (sx, _) in ((LM, LS), (RM, RS)):
        cx = x + w / 2
        for a, b in (("id", "scr"), ("scr", "sought"), ("sought", "assess")):
            arrow((cx, ROW[a][0]), (cx, ROW[b][0] + ROW[b][1]))
        arrow((cx, ROW["assess"][0]), (cx, iy + ih))
        for r in ("id", "scr", "sought", "assess"):
            y, h = ROW[r]
            arrow((x + w, y + h / 2), (sx, y + h / 2))

    for ext in ("png", "pdf"):
        fig.savefig(os.path.join(FIG, f"Fig1_PRISMA.{ext}"))
    plt.close(fig)
    print(f"  Fig1_PRISMA.png / .pdf   (PRISMA 2020 템플릿 · 공통 글자 {f:.2f} pt · 최종 n = {D['n_included']})")


if __name__ == "__main__":
    main()
