# -*- coding: utf-8 -*-
"""
Paper32 — 논문 Figure 생성 (투고용 영문 라벨, 300dpi)
Fig1 PRISMA flow · Fig2 Forest plots (4 clusters) · Fig3 Evidence map heatmap
· Fig4 Direction by domain · Fig5 Measurement generation over time
출력: figures/Fig1_prisma.png ... (+ .pdf)
"""
import sys, os, csv, math
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FIG = os.path.join(BASE, "figures")
os.makedirs(FIG, exist_ok=True)
import viz_theme as T
T.apply()
INK = T.INK; ACC = T.BLUE; MUT = T.AXIS; POS = T.BLUE; NEG = T.ORANGE


def save(fig, name):
    for ext in ("png", "pdf"):
        fig.savefig(os.path.join(FIG, f"{name}.{ext}"))
    plt.close(fig)
    print(f"  {name}.png / .pdf")


# ---------------- Fig 1: PRISMA ----------------
def fig_prisma():
    """좌: 주 흐름(폭 4.6) · 우: 배제 사유(폭 4.4) · 각 단계 높이를 내용에 맞춰 배치."""
    fig, ax = plt.subplots(figsize=(8.4, 8.0))
    ax.set_xlim(0, 10.4); ax.set_ylim(0, 10.6); ax.axis("off")
    LX, LW, RX, RW = 0.95, 4.55, 5.85, 4.35
    CX = LX + LW / 2

    def box(x, y, w, h, text, fc="white", ec=INK, fs=8.0, bold_last=False):
        ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle="round,pad=0.05,rounding_size=0.08",
                                    fc=fc, ec=ec, lw=0.9))
        ax.text(x + w / 2, y + h / 2, text, ha="center", va="center", fontsize=fs,
                color=INK, linespacing=1.45)

    def arrow(x1, y1, x2, y2):
        ax.add_patch(FancyArrowPatch((x1, y1), (x2, y2), arrowstyle="-|>", mutation_scale=10,
                                     lw=0.9, color=INK, shrinkA=0, shrinkB=1))

    for y0, y1, lab in [(8.55, 10.45, "Identification"), (6.05, 8.35, "Screening"),
                        (2.15, 5.85, "Eligibility"), (0.15, 1.95, "Included")]:
        ax.add_patch(FancyBboxPatch((0.06, y0), 0.5, y1 - y0, boxstyle="round,pad=0.02",
                                    fc="#eef2f5", ec="none"))
        ax.text(0.31, (y0 + y1) / 2, lab, rotation=90, ha="center", va="center",
                fontsize=9, color=ACC, fontweight="bold")

    # Identification
    box(LX, 9.35, LW, 1.05,
        "Records identified from databases (2 Aug 2026)\nWeb of Science 1,010 · Scopus 850 · PubMed 213\n"
        r"$\bf{n\ =\ 2{,}073}$")
    box(RX, 9.5, RW, 0.75, "Duplicates removed\nn = 757", fc="#f6f6f6", ec=MUT, fs=7.8)
    arrow(LX + LW, 9.875, RX, 9.875)
    arrow(CX, 9.35, CX, 8.35)

    # Screening
    box(LX, 7.55, LW, 0.8, "Records screened (title/abstract)\n" + r"$\bf{n\ =\ 1{,}316}$")
    box(RX, 6.75, RW, 1.85,
        r"$\bf{Records\ excluded\ (n = 1{,}127)}$" + "\n"
        "Animal / wildlife (soundscape ecology)   479\n"
        "Perception, health or physiology only   325\n"
        "Setting not eligible   133\nNot empirical   109\nNo acoustic variable   81",
        fc="#f6f6f6", ec=MUT, fs=7.4)
    arrow(LX + LW, 7.95, RX, 7.95)
    arrow(CX, 7.55, CX, 6.65)

    # Eligibility
    box(LX, 5.85, LW, 0.8, "Reports sought for retrieval\n" + r"$\bf{n\ =\ 189}$")
    box(RX, 5.15, RW, 1.4,
        r"$\bf{Reports\ not\ retrieved\ (n = 89)}$" + "\n"
        "No institutional access   84\nPay-per-view only   3\nAbstract only   2",
        fc="#f6f6f6", ec=MUT, fs=7.4)
    arrow(LX + LW, 6.25, RX, 6.25)
    arrow(CX, 5.85, CX, 4.95)

    box(LX, 4.15, LW, 0.8, "Reports assessed for eligibility (full text)\n" + r"$\bf{n\ =\ 100}$")
    box(RX, 3.25, RW, 1.65,
        r"$\bf{Reports\ excluded\ (n = 16)}$" + "\n"
        "Setting not eligible   7\nNo observable behavioural outcome   6\n"
        "Perceptual outcome only   1\nNo acoustic exposure   1\nNot empirical   1",
        fc="#f6f6f6", ec=MUT, fs=7.4)
    arrow(LX + LW, 4.55, RX, 4.55)
    box(RX, 2.25, RW, 0.8,
        "Reserved for sensitivity analysis\n(behavioural-intention outcomes)   n = 3",
        fc="#fbf7ec", ec=MUT, fs=7.4)
    arrow(LX + LW, 4.25, RX, 2.65)
    arrow(CX, 4.15, CX, 1.95)

    # Included
    box(LX, 1.05, LW, 0.85,
        "Studies included in qualitative synthesis\n" + r"$\bf{n\ =\ 81}$", fc="#eaf3ee", ec=POS)
    ax.add_patch(FancyBboxPatch((LX, 0.15), LW, 0.75, boxstyle="round,pad=0.05,rounding_size=0.08",
                                fc="#eaf3ee", ec=POS, lw=0.9))
    ax.text(LX + LW / 2, 0.66, "Contributing to meta-analysis:  n = 15",
            ha="center", va="center", fontsize=7.8, color=INK, fontweight="bold")
    ax.text(LX + LW / 2, 0.40, "walking speed 4 · staying 3 · social 3 · correlation 5",
            ha="center", va="center", fontsize=7.5, color=INK)
    arrow(CX, 1.05, CX, 0.9)
    save(fig, "Fig1_PRISMA")


# ---------------- Fig 2: Forest plots ----------------
def forest_panel(ax, rows, pooled, title, xlab, xlim):
    """라벨은 y축 tick(플롯 바깥)에 배치 — 데이터와 겹치지 않게."""
    ys = list(range(len(rows), 0, -1))
    for y, r in zip(ys, rows):
        lo, hi = r["est"] - 1.96 * math.sqrt(r["var"]), r["est"] + 1.96 * math.sqrt(r["var"])
        ax.plot([lo, hi], [y, y], color=INK, lw=1.0)
        ax.plot([lo, lo], [y - .1, y + .1], color=INK, lw=1.0)
        ax.plot([hi, hi], [y - .1, y + .1], color=INK, lw=1.0)
        size = 26 + 240 * (1 / r["var"]) / max(1 / rr["var"] for rr in rows)
        ax.scatter([r["est"]], [y], s=size, marker="s",
                   color=T.BLUE, zorder=3, edgecolor="white", lw=0.9)
    # diamond
    y0 = 0.35
    e, lo, hi = pooled["est"], pooled["lo"], pooled["hi"]
    ax.add_patch(plt.Polygon([[lo, y0], [e, y0 + .24], [hi, y0], [e, y0 - .24]],
                             closed=True, fc=T.FRAME_POS, ec=T.FRAME_POS, alpha=.95, zorder=4))
    ax.axvline(0, color=MUT, lw=0.8, ls="--", zorder=0)
    ax.set_xlim(*xlim); ax.set_ylim(-0.25, len(rows) + 0.75)
    ax.set_yticks(ys + [y0])
    ax.set_yticklabels([r["label"] for r in rows] +
                       [f"Pooled ({pooled['model']}), k = {pooled['k']}"], fontsize=7.3)
    for lbl in ax.get_yticklabels()[-1:]:
        lbl.set_color(ACC); lbl.set_fontweight("bold")
    ax.tick_params(axis="y", length=0)
    ax.set_xlabel(xlab, fontsize=8)
    ax.set_title(title, fontsize=9, loc="left", pad=6)
    ax.text(0.99, 0.02, f"$I^2$ = {pooled['I2']:.0f}%   p = {pooled['p']:.3f}",
            transform=ax.transAxes, ha="right", fontsize=7.4, color=MUT)
    for s in ("top", "right", "left"):
        ax.spines[s].set_visible(False)


def fig_forest():
    walk = [dict(label="Franěk 2018 · birdsong vs traffic", est=-1.024, var=0.283 ** 2),
            dict(label="Franěk 2019 Exp1 · birdsong vs city noise", est=-0.359, var=0.265 ** 2),
            dict(label="Franěk 2019 Exp2 · birdsong vs city noise", est=-0.896, var=0.316 ** 2),
            dict(label="Oases street · nature vs traffic (field obs.)", est=0.227, var=0.273 ** 2)]
    w_p = dict(est=-0.500, lo=-1.411, hi=0.412, k=4, I2=75.7, p=0.179, model="REML+HK")
    stay = [dict(label="Aletta 2016 · music vs no music", est=0.388, var=0.0106),
            dict(label="Ba & Kang 2020 · music vs no sound", est=0.614, var=0.0432),
            dict(label="Bao 2026 · natural sound index (long stay)", est=0.214, var=0.0053)]
    s_p = dict(est=0.313, lo=-0.076, hi=0.703, k=3, I2=54.4, p=0.074, model="REML+HK")
    # ★ 인용추적으로 추가된 효과는 ▲ 표시(citation searching). 수치 = ma/ma_v2_summary.md
    soc = [dict(label="Mathews & Canon 1975 · quiet vs mower noise (helping) ▲",
                est=1.073, var=0.0999),
           dict(label="Chen 2023 · natural vs noise (group interaction)", est=0.977, var=0.0613),
           dict(label="Chen 2024 · natural vs noise (paired interaction)", est=0.547, var=0.0284),
           dict(label="Moser 1988 · quiet vs roadworks (helping)", est=0.432, var=0.0232)]
    so_p = dict(est=0.646, lo=0.188, hi=1.104, k=4, I2=48.6, p=0.021, model="REML+HK")
    corr = [dict(label="Montes González 2022 · LAeq ↔ speech disruption ▲",
                 est=np.arctanh(0.650), var=1 / (29 - 3)),
            dict(label="Xu 2024 · pleasantness ↔ static behaviour", est=np.arctanh(0.564), var=1 / (419 - 3)),
            dict(label="Bao 2023 · dwell time ↔ restorativeness", est=np.arctanh(0.551), var=1 / (180 - 3)),
            dict(label="Béjaïa 2025 · sound ↔ walking comfort", est=np.arctanh(0.400), var=1 / (58 - 3)),
            dict(label="Study 1018 · sitting/walking groups", est=np.arctanh(0.360), var=1 / (310 - 3)),
            dict(label="Wang 2025 · natural events ↔ queuing time", est=np.arctanh(0.209), var=1 / (315 - 3)),
            dict(label="Cao & Kang 2021 · companionship → sound noticing ▲",
                 est=np.arctanh(0.165), var=1 / (301 - 3))]
    c_p = dict(est=0.435, lo=0.231, hi=0.639, k=7, I2=90.5, p=0.002, model="REML+HK")

    fig, axes = plt.subplots(2, 2, figsize=(12.6, 6.4))
    forest_panel(axes[0, 0], walk, w_p, "(a) Walking speed — natural sound vs anthropogenic noise",
                 "Hedges' g  (negative = faster walking under noise)", (-2.2, 1.4))
    forest_panel(axes[0, 1], stay, s_p, "(b) Staying / dwell time — positive sound vs control",
                 "Hedges' g  (positive = longer stay)", (-0.8, 1.4))
    forest_panel(axes[1, 0], soc, so_p, "(c) Social interaction — natural/quiet vs noise",
                 "Hedges' g  (positive = more interaction)", (-0.6, 2.1))
    forest_panel(axes[1, 1], corr, c_p, "(d) Soundscape perception ↔ behaviour (correlational)",
                 "Fisher's z  (back-transformed r = 0.41)", (-0.2, 1.4))
    fig.suptitle("Meta-analytic effects across four behavioural clusters", fontsize=11, x=0.007, ha="left")
    fig.text(0.993, 0.982, "▲ = study added by citation searching", fontsize=7.6, color=MUT, ha="right")
    fig.tight_layout(rect=[0, 0, 1, 0.955], w_pad=3.2, h_pad=2.4)
    save(fig, "Fig2_Forest")


# ---------------- Fig 3: Evidence map ----------------
DOM_KEY = ["movement", "staying", "space-use", "social", "activity"]
DOM_EN = ["Movement", "Staying", "Space use", "Social", "Activity"]
SRC_KEY = ["교통/도로소음", "일반 소음", "자연음", "음악/부가음", "인간·군중음", "항공기소음"]
SRC_EN = ["Traffic/road", "General noise", "Natural sounds", "Music/added",
          "Human/crowd", "Aircraft"]


def load_counts(kind):
    """증거지도 집계는 하드코딩하지 않고 rebuild_evidence_map.py 산출에서 읽는다."""
    out = {}
    with open(os.path.join(BASE, "fulltext", "evidence_counts_v2.csv"),
              encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            if r["kind"] == kind:
                out[(r["row"], r["col"])] = int(r["n"])
    return out


def fig_evidence_map():
    doms, srcs = DOM_EN, SRC_EN
    c = load_counts("domain_x_source")
    M = np.array([[c.get((d, s), 0) for s in SRC_KEY] for d in DOM_KEY])
    fig, ax = plt.subplots(figsize=(6.4, 6.4))
    im = ax.imshow(M, cmap=T.SEQ, vmin=0, vmax=M.max(), aspect="auto")
    ax.set_xticks(range(len(srcs))); ax.set_xticklabels(srcs, fontsize=8, rotation=18, ha="right")
    ax.set_yticks(range(len(doms))); ax.set_yticklabels(doms, fontsize=8.5)
    for i in range(M.shape[0]):
        for j in range(M.shape[1]):
            ax.text(j, i, M[i, j], ha="center", va="center", fontsize=8.5,
                    color="white" if M[i, j] >= M.max() * 0.42 else INK,
                    fontweight="bold" if M[i, j] == 0 else "normal")
    ax.set_xticks(np.arange(-.5, len(srcs), 1), minor=True)
    ax.set_yticks(np.arange(-.5, len(doms), 1), minor=True)
    ax.grid(which="minor", color=T.SURF, linewidth=2.2)
    ax.tick_params(which="minor", length=0)
    for sp in ax.spines.values(): sp.set_visible(False)
    ax.set_title("Evidence map: number of included studies by behavioural domain × sound source",
                 fontsize=9.5, loc="left", pad=8)
    cb = fig.colorbar(im, ax=ax, shrink=0.82, pad=0.02); cb.ax.tick_params(labelsize=7)
    cb.set_label("studies (n)", fontsize=7.5)
    fig.text(0.012, 0.012,
             "Aircraft noise is effectively absent (n = 2) despite an established health literature — the clearest evidence gap.",
             fontsize=7.2, color=MUT, ha="left")
    fig.tight_layout(rect=[0, 0.055, 1, 1])
    save(fig, "Fig3_EvidenceMap")


# ---------------- Fig 4: Direction × domain ----------------
def fig_direction():
    doms = DOM_EN
    c = load_counts("direction_x_domain")
    fwd = np.array([c.get((d, "forward"), 0) for d in DOM_KEY])
    rev = np.array([c.get((d, "reverse"), 0) for d in DOM_KEY])
    both = np.array([c.get((d, "both"), 0) for d in DOM_KEY])
    y = np.arange(len(doms)); h = 0.72
    fig, ax = plt.subplots(figsize=(6.4, 6.4))
    ax.barh(y, fwd, h, label="Forward (sound → behaviour)", color=T.BLUE, edgecolor=T.SURF, lw=1.4)
    ax.barh(y, rev, h, left=fwd, label="Reverse (behaviour → soundscape)", color=T.ORANGE, edgecolor=T.SURF, lw=1.4)
    ax.barh(y, both, h, left=fwd + rev, label="Bidirectional", color=T.MUTED, edgecolor=T.SURF, lw=1.4)
    for i in range(len(doms)):
        ax.text(fwd[i] / 2, i, str(fwd[i]), ha="center", va="center", color="white", fontsize=7.6)
        ax.text(fwd[i] + rev[i] / 2, i, str(rev[i]), ha="center", va="center", color="white", fontsize=7.6)
    ax.set_yticks(y); ax.set_yticklabels(doms, fontsize=8.5); ax.invert_yaxis()
    ax.set_xlabel("Number of extracted study–outcome records", fontsize=8)
    top=int((fwd+rev+both).max()); ax.set_xticks(range(0, top+11, 10)); ax.set_axisbelow(True)
    ax.xaxis.grid(True, color=T.GRID, lw=0.7); ax.yaxis.grid(False)
    ax.set_title("Direction of the studied relationship by behavioural domain", fontsize=9.5, loc="left", pad=26)
    ax.legend(fontsize=7.4, frameon=False, ncol=3, loc="lower center", bbox_to_anchor=(0.5, 1.005), columnspacing=1.6, handlelength=1.4, handleheight=0.9)
    for s in ("top", "right", "left"): ax.spines[s].set_visible(False)
    save(fig, "Fig4_Direction")


# ---------------- Fig 5: Measurement generations ----------------
def fig_methods():
    """측정세대 카운트는 하드코딩하지 않고 table1의 measure_gen에서 산출한다.
    코딩 규칙(Table 1 각주와 동일): 한 연구가 두 세대를 함께 쓰면 둘 다 계상(다중코딩).
    노출 계측 전용 장비(소음계 등)는 '행태' 측정이 아니므로 G3로 세지 않는다."""
    from collections import Counter, defaultdict
    per = defaultdict(Counter)
    with open(os.path.join(BASE, "fulltext", "table1_v2.csv"),
              encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            y = int(r["year"]) if r["year"].isdigit() else 0
            band = 0 if y < 2010 else (1 if y < 2020 else 2)
            for g in (x.strip() for x in r["measure_gen"].split(";")):
                if g in ("G1", "G2", "G3"):
                    per[band][g] += 1
    bands = ["≤2009", "2010–2019", "2020–2026"]
    sr = [per[i]["G1"] for i in range(3)]
    ob = [per[i]["G2"] for i in range(3)]
    se = [per[i]["G3"] for i in range(3)]
    x = np.arange(len(bands)); w = 0.235
    fig, ax = plt.subplots(figsize=(6.4, 6.4))
    ax.bar(x - w, sr, w, label="Self-report", color=T.MUTED, edgecolor=T.SURF, lw=1.2)
    ax.bar(x, ob, w, label="On-site observation", color=T.BLUE, edgecolor=T.SURF, lw=1.2)
    ax.bar(x + w, se, w, label="Sensor / GPS / video / big data", color=T.ORANGE, edgecolor=T.SURF, lw=1.2)
    for xi, vals in zip(x, zip(sr, ob, se)):
        for dx, v in zip((-w, 0, w), vals):
            if v: ax.text(xi + dx, v + 0.5, str(v), ha="center", fontsize=7.4, color=INK)
    ax.set_xticks(x); ax.set_xticklabels(bands, fontsize=8.5)
    ax.set_ylabel("Studies (n)", fontsize=8)
    top = max(sr + ob + se)
    ax.set_yticks(range(0, top + 6, 5)); ax.set_ylim(0, top + 4)
    ax.set_axisbelow(True); ax.yaxis.grid(True, color=T.GRID, lw=0.7)
    ax.xaxis.grid(False)
    ax.set_title("Behavioural measurement methods over time (included studies)",
                 fontsize=9.5, loc="left", pad=20)
    ax.text(0.0, 1.012, "studies using two generations are counted in both",
            transform=ax.transAxes, ha="left", fontsize=7.0, color=T.AXIS)
    ax.legend(fontsize=7.4, frameon=False)
    for s in ("top", "right"): ax.spines[s].set_visible(False)
    save(fig, "Fig5_Methods")


if __name__ == "__main__":
    print("Figures →", FIG)
    # ⚠️ fig_prisma()는 DB 검색 단일 갈래판(구버전). 인용추적 갈래를 포함한 정본 Fig1은
    #    make_fig1_prisma_v2.py 가 생성한다 — 여기서 호출하면 정본을 덮어쓰므로 비활성.
    fig_forest(); fig_evidence_map(); fig_direction(); fig_methods()
    print("완료 (Fig1은 make_fig1_prisma_v2.py 로 별도 생성)")
