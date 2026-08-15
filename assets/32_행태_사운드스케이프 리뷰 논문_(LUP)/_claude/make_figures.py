# -*- coding: utf-8 -*-
"""
Paper32 — 논문 Figure 2~5 (투고용 영문 라벨, 300 dpi)

★ 이 파일에는 **데이터 수치를 절대 두지 않는다.** 전부 `figures/fig_data.json`
  (= `fig_data.py` 산출)에서 읽는다. 종전에는 효과값·집계값을 하드코딩해 두어
  정본이 갱신돼도 그림이 옛 숫자를 그렸다(Fig 5 는 102편, Fig 6 은 81편 기준이었다).

판형: Elsevier 2단 폭 190 mm ≈ 7.48 in 기준. 종횡비는 내용에 맞춰 정한다.
출력: figures/Fig2_Forest · Fig3_EvidenceMap · Fig4_Direction · Fig5_Methods (png+pdf)
"""
import sys, os, json, math
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FIG = os.path.join(BASE, "figures")
os.makedirs(FIG, exist_ok=True)
sys.path.insert(0, BASE)
import viz_theme as T
T.apply()
INK, ACC, MUT = T.INK, T.BLUE, T.AXIS

W2 = T.W_FULL      # ★ 작화 폭 = docx 삽입 폭 — 그린 글자 크기가 곧 인쇄 크기
D = json.load(open(os.path.join(FIG, "fig_data.json"), encoding="utf-8"))


def save(fig, name):
    for ext in ("png", "pdf"):
        fig.savefig(os.path.join(FIG, f"{name}.{ext}"))
    plt.close(fig)
    print(f"  {name}.png / .pdf")


# ═══════════════════ Fig 2: Forest ═══════════════════
FOREST_LABEL = {
    "461": "Franěk 2018 · birdsong vs traffic",
    "532": "Franěk 2019 · birdsong vs city noise (Exp 1–2)",
    "617": "Berkouk 2020 · nature vs traffic",
    "323": "Aletta 2016 · music vs no music",
    "665": "Ba & Kang 2020 · music vs no sound",
    "1280": "Fu 2026 · natural sound index",
    "931": "Chen 2023 · natural vs noise (group)",
    "1069": "Chen 2024 · natural vs noise (paired)",
    "14": "Moser 1988 · quiet vs roadworks",
    "CT0025": "Mathews & Canon 1975 · quiet vs mower",
    "1076": "Guo 2024 · pleasantness ↔ static behaviour",
    "1221": "Zhou 2026 · natural events ↔ queuing",
    "1177": "Mansouri 2025 · sound ↔ walking comfort",
    "980": "Bao 2023 · dwell time ↔ restorativeness",
    "CT0126": "Montes González 2023 · LAeq ↔ speech",
    "CT0184": "Cao & Kang 2021 · companionship → noticing",
}
PANE = [("walking", "(a) Walking speed", "natural sound vs anthropogenic noise",
         "Hedges' $g$   (negative = faster under noise)"),
        ("staying", "(b) Staying / dwell time", "positive sound vs control",
         "Hedges' $g$   (positive = longer stay)"),
        ("social", "(c) Social interaction", "natural or quiet vs noise",
         "Hedges' $g$   (positive = more interaction)"),
        ("correlation", "(d) Perception ↔ behaviour", "correlational",
         "Fisher's $z$   (positive = appraisal tracks behaviour)")]
# (d) 는 back-transformed r 까지 붙어 x라벨이 길어지므로 축약형을 쓴다


def forest_panel(ax, key, head, sub, xlab):
    blk = D["ma"][key]
    eff = blk["effects"]; pl = blk["pooled"]
    n = len(eff)
    # 각 효과는 y = n..1, 풀링 다이아몬드는 y = -0.35 (구분선 아래)
    ys = list(range(n, 0, -1))
    inv = [1 / e["var"] for e in eff]
    for y, e, w in zip(ys, eff, inv):
        se = math.sqrt(e["var"])
        lo, hi = e["est"] - 1.96 * se, e["est"] + 1.96 * se
        ax.plot([lo, hi], [y, y], color=INK, lw=0.9, solid_capstyle="butt", zorder=2)
        for x in (lo, hi):
            ax.plot([x, x], [y - .13, y + .13], color=INK, lw=0.9, zorder=2)
        ax.scatter([e["est"]], [y], s=22 + 150 * w / max(inv), marker="s",
                   color=T.BLUE, zorder=3, edgecolor="white", lw=0.8)
    yD = -0.45
    ax.axhline(0.28, color=T.GRID, lw=0.8, zorder=1)
    ax.add_patch(plt.Polygon([[pl["lo"], yD], [pl["est"], yD + .26],
                              [pl["hi"], yD], [pl["est"], yD - .26]],
                             closed=True, fc=T.FRAME_POS, ec=T.FRAME_POS, zorder=4))
    ax.axvline(0, color=MUT, lw=0.8, ls=(0, (4, 3)), zorder=0)

    # x 범위: 효과 CI 와 풀링 CI 를 모두 담고 6% 여유
    xs = [pl["lo"], pl["hi"]]
    for e in eff:
        se = math.sqrt(e["var"])
        xs += [e["est"] - 1.96 * se, e["est"] + 1.96 * se]
    lo, hi = min(xs), max(xs)
    pad = (hi - lo) * 0.10
    ax.set_xlim(lo - pad, hi + pad)
    ax.set_ylim(yD - 0.95, n + 0.55)
    ax.set_yticks(ys + [yD])
    lbl = []
    for e in eff:
        s = FOREST_LABEL.get(str(e["uid"]), str(e["uid"]))
        lbl.append(s + ("  ▲" if e["route"] == "citation-tracking" else ""))
    ax.set_yticklabels(lbl + [f"Pooled,  $k$ = {pl['k']}"], fontsize=7.5)
    ax.get_yticklabels()[-1].set_color(T.FRAME_POS)
    ax.get_yticklabels()[-1].set_fontweight("bold")
    ax.tick_params(axis="y", length=0, pad=2)
    ax.tick_params(axis="x", labelsize=7.5)
    ax.xaxis.set_major_locator(plt.MaxNLocator(5))
    # 통계는 패널 오른쪽 위에 별도 텍스트로 — x축 라벨과 뒤섞지 않는다(가독성)
    r_txt = ""
    if blk.get("back_r"):
        r_txt = (f"\n$r$ = {math.tanh(pl['est']):+.2f} "
                 f"[{math.tanh(pl['lo']):+.2f}, {math.tanh(pl['hi']):+.2f}]")
        xlab = "Fisher's $z$"
    ax.text(1.0, 1.02, f"$I^2$ = {pl['I2']:.0f}%,  $p$ = {pl['p']:.3f}{r_txt}",
            transform=ax.transAxes, ha="right", va="bottom", fontsize=7.2,
            color=T.INK2, linespacing=1.4)
    ax.set_xlabel(xlab, fontsize=7.8, labelpad=3)
    ax.set_title(head, fontsize=9.0, loc="left", pad=14,
                 fontweight="bold")   # 부제는 캡션이 담당
    for s in ("top", "right", "left"):
        ax.spines[s].set_visible(False)


def fig_forest():
    """4행 1열. 2×2 로는 패널 폭이 3.4 in 뿐이라 라벨이 플롯을 잡아먹는다(실측).
    한 열로 세우면 각 패널이 전체 폭을 쓰고, 라벨 왼쪽 여백을 고정해 네 패널의
    x=0 기준선이 세로로 정렬된다 — 클러스터 간 비교가 가능해진다."""
    ns = [len(D["ma"][k]["effects"]) for k, *_ in PANE]
    fig, axes = plt.subplots(len(PANE), 1, figsize=(W2, 6.9),
                             gridspec_kw={"height_ratios": [n + 2.4 for n in ns]})
    for (key, head, sub, xlab), ax in zip(PANE, axes):
        forest_panel(ax, key, head, sub, xlab)
    # ★ 제목·부제를 그림에 넣지 않는다 — 캡션이 담당한다(저널 관행).
    fig.subplots_adjust(left=0.335, right=0.985, top=0.955, bottom=0.055, hspace=1.15)
    save(fig, "Fig2_Forest")


# ═══════════════════ Fig 3: Evidence map ═══════════════════
def fig_evidence_map():
    em = D["evidence_map"]
    M = np.array(em["cells"], float)
    fig, ax = plt.subplots(figsize=(W2, 2.75))
    norm = plt.Normalize(vmin=0, vmax=M.max())
    ax.imshow(M, cmap=T.SEQ, aspect="auto", norm=norm)
    ax.set_xticks(range(len(em["cols"]))); ax.set_xticklabels(em["cols"], fontsize=8)
    ax.set_yticks(range(len(em["rows"]))); ax.set_yticklabels(em["rows"], fontsize=8)
    ax.tick_params(length=0)
    ax.set_xticks(np.arange(-.5, len(em["cols"]), 1), minor=True)
    ax.set_yticks(np.arange(-.5, len(em["rows"]), 1), minor=True)
    ax.grid(which="minor", color="white", lw=1.5)
    ax.tick_params(which="minor", length=0)
    for s in ax.spines.values():
        s.set_visible(False)
    # 컬러바는 뺐다 — 모든 칸에 정확한 값이 인쇄돼 있어 중복 장식이다.
    for i in range(M.shape[0]):
        for j in range(M.shape[1]):
            v = int(M[i, j])
            ax.text(j, i, str(v), ha="center", va="center", fontsize=8.5,
                    color=T.ink_on(T.SEQ(norm(M[i, j]))),
                    fontweight=("bold" if v == 0 else "normal"))
    fig.tight_layout()
    save(fig, "Fig3_EvidenceMap")


# ═══════════════════ Fig 4: Direction ═══════════════════
def fig_direction():
    d = D["direction"]
    y = np.arange(len(d["rows"]))[::-1]
    fig, ax = plt.subplots(figsize=(W2, 2.65))
    h = 0.62
    SEG = [("Forward — sound → behaviour", d["forward"], T.BLUE),
           ("Bidirectional", d["both"], T.NEUT),
           ("Reverse — behaviour → sound", d["reverse"], T.TERRA)]
    left = np.zeros(len(y))
    for lab, vals, col in SEG:
        ax.barh(y, vals, h, left=left, color=col, label=lab,
                edgecolor=T.SURF, linewidth=0.7)
        for yi, l0, v in zip(y, left, vals):
            if v >= 3:
                ax.text(l0 + v / 2, yi, str(v), ha="center", va="center",
                        fontsize=7.8, color=T.ink_on(col), fontweight="bold")
        left += np.array(vals, float)
    for yi, f, b, r in zip(y, d["forward"], d["both"], d["reverse"]):
        ax.text(f + b + r + 1.2, yi, f"{r / (f + r) * 100:.0f}% reverse",
                va="center", fontsize=7.5, color=T.INK2)
    ax.set_yticks(y); ax.set_yticklabels(d["rows"], fontsize=8.5)
    ax.set_xlim(0, max(a + b + c for a, b, c in
                       zip(d["forward"], d["both"], d["reverse"])) * 1.22)
    ax.set_xlabel("study × behavioural-domain records", fontsize=8)
    ax.tick_params(length=0)
    ax.xaxis.set_visible(False)               # 값이 전부 인쇄돼 있어 눈금이 중복이다
    for s in ("top", "right", "left", "bottom"):
        ax.spines[s].set_visible(False)
    ax.legend(loc="lower left", bbox_to_anchor=(0.0, 1.0), ncol=3, fontsize=7.5,
              handlelength=1.1, columnspacing=1.5, handletextpad=0.5)
    fig.tight_layout()
    save(fig, "Fig4_Direction")


# ═══════════════════ Fig 5: Measurement generations ═══════════════════
def fig_methods():
    g = D["generations"]
    x = np.arange(len(g["bands"])); w = 0.26
    # 세대는 순서형 — 3색이 아니라 한 hue 의 3단 램프(밝음→어두움 = G1→G3)
    SER = [("G1  Self-report", g["G1"], T.ORD3[0]),
           ("G2  Systematic observation", g["G2"], T.ORD3[1]),
           ("G3  Sensing · GPS · video", g["G3"], T.ORD3[2])]
    fig, ax = plt.subplots(figsize=(W2, 2.65))
    for i, (lab, vals, col) in enumerate(SER):
        pos = x + (i - 1) * w
        ax.bar(pos, vals, w * 0.9, color=col, label=lab)
        for p, v in zip(pos, vals):
            if v:
                ax.text(p, v + 0.8, str(v), ha="center", fontsize=7.8, color=INK)
    ax.set_xticks(x); ax.set_xticklabels(g["bands"], fontsize=8.5)
    ax.set_ylim(0, max(max(v) for _, v, _ in SER) * 1.20)
    ax.set_ylabel("studies (n)", fontsize=8)
    ax.tick_params(length=0)
    ax.grid(axis="y"); ax.set_axisbelow(True)
    for s in ("top", "right", "left"):
        ax.spines[s].set_visible(False)
    ax.legend(loc="upper left", fontsize=7.5, handlelength=1.1, handletextpad=0.5)
    ax.margins(x=0.06)
    fig.tight_layout()
    save(fig, "Fig5_Methods")


if __name__ == "__main__":
    print(f"Figures → {FIG}")
    fig_forest(); fig_evidence_map(); fig_direction(); fig_methods()
    print("완료 (Fig1·6·7·8 은 각 전용 스크립트)")
