# -*- coding: utf-8 -*-
"""
Paper32 — Figure 8: 방법론적 질 평가 (MMAT 2018) — PRISMA 2020 item 21 이행
(a) MMAT 범주별 등급 분포
(b) 문항별 판정 비율 (Yes / Can't tell / No)
근거: fulltext/quality_all.csv · quality_detail_all.csv
출력: figures/Fig8_Quality.png|pdf
"""
import sys, os, csv
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import viz_theme as T

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
FIG = os.path.join(BASE, "figures")
os.makedirs(FIG, exist_ok=True)
T.apply()

YES, CT, NO = T.BLUE, T.MUTED, T.ORANGE

CATS = [("1", "Qualitative"), ("2", "Quantitative RCT"),
        ("3", "Quantitative non-randomised"), ("4", "Quantitative descriptive"),
        ("5", "Mixed methods")]
ITEMS = {
    "1.1": "approach appropriate to question", "1.2": "data collection adequate",
    "1.3": "findings derived from data", "1.4": "interpretation substantiated",
    "1.5": "coherence across components",
    "2.1": "randomisation appropriate", "2.2": "groups comparable at baseline",
    "2.3": "complete outcome data", "2.4": "assessors blinded",
    "2.5": "adherence to assigned condition",
    "3.1": "participants representative", "3.2": "measurements appropriate",
    "3.3": "complete outcome data", "3.4": "confounders accounted for",
    "3.5": "exposure occurred as intended",
    "4.1": "sampling strategy relevant", "4.2": "sample representative",
    "4.3": "measurements appropriate", "4.4": "nonresponse bias low",
    "4.5": "statistical analysis appropriate",
    "5.1": "rationale for mixed design", "5.2": "components integrated",
    "5.3": "integration outputs interpreted", "5.4": "divergences addressed",
    "5.5": "components meet each tradition",
}


def main():
    # ⚠️ 종전 코드는 102편(민감도 4편 포함) 전건을 집계해 본문 §3.3(98편)과 문항 백분율이
    #    어긋났다(4.4 = 41% vs 31%). FINAL_INCLUDE 로 맞춘다.
    verd = {r["uid"]: r["final_verdict"] for r in csv.DictReader(
        open(os.path.join(FT, "corpus_v4_verdicts.csv"), encoding="utf-8-sig"))}
    keep = lambda u: verd.get(u) == "FINAL_INCLUDE"
    q = [r for r in csv.DictReader(open(os.path.join(FT, "quality_v2.csv"), encoding="utf-8-sig"))
         if keep(r["uid"])]
    d = [r for r in csv.DictReader(open(os.path.join(FT, "quality_detail_v2.csv"),
                                        encoding="utf-8-sig")) if keep(r["uid"])]

    from collections import Counter, defaultdict
    # 범주 라벨 → 숫자 코드
    def catcode(s):
        s = (s or "").lower()
        if "mixed" in s: return "5"
        if "descriptive" in s: return "4"
        if "non-random" in s or "non_random" in s: return "3"
        if "rct" in s or "random" in s: return "2"
        if "qualitative" in s: return "1"
        return "?"

    tier_by_cat = defaultdict(Counter)
    for r in q:
        tier_by_cat[catcode(r["mmat_category"])][r["quality_tier"]] += 1

    item_v = defaultdict(Counter)
    for r in d:
        i = (r["item_no"] or "").strip()
        if i in ITEMS:
            item_v[i][(r["verdict"] or "").strip().upper()] += 1

    fig, axes = plt.subplots(2, 1, figsize=(8.4, 6.6),
                             gridspec_kw={"height_ratios": [1.0, 1.55], "hspace": 0.42})

    # ── (a) 범주별 등급 ─────────────────────────────────────────────
    ax = axes[0]
    codes = [c for c, _ in CATS if tier_by_cat.get(c)]
    labels = [f"{dict(CATS)[c]}" for c in codes]
    tiers = [("high", T.SEQ(0.85), "high (4–5 criteria met)"),
             ("moderate", T.SEQ(0.5), "moderate (3)"),
             ("low", T.SEQ(0.18), "low (0–2)")]
    y = np.arange(len(codes))[::-1].astype(float)
    left = np.zeros(len(codes))
    for key, col, lab in tiers:
        v = np.array([tier_by_cat[c].get(key, 0) for c in codes], float)
        ax.barh(y, v, left=left, height=0.58, color=col, edgecolor=T.SURF,
                linewidth=0.8, label=lab, zorder=3)
        for yi, (l0, vi) in enumerate(zip(left, v)):
            if vi >= 2:
                ax.text(l0 + vi / 2, y[yi], f"{int(vi)}", ha="center", va="center",
                        fontsize=7.2, color=T.SURF if key != "low" else T.INK,
                        fontweight="bold", zorder=5)
        left += v
    totals = left
    for yi, tot in zip(y, totals):
        ax.text(tot + 0.7, yi, f"n = {int(tot)}", va="center", fontsize=7.4, color=T.INK2)
    ax.set_yticks(y); ax.set_yticklabels(labels, fontsize=8.2)
    ax.set_xlabel("studies (n)", fontsize=8)
    ax.set_xlim(0, max(totals) * 1.16)
    ax.grid(axis="x", zorder=0); ax.set_axisbelow(True)
    for s in ("top", "right", "left"):
        ax.spines[s].set_visible(False)
    ax.legend(fontsize=7.2, loc="upper right", ncol=1, handlelength=1.0,
              columnspacing=1.1, borderpad=0.2, labelspacing=0.34)
    ax.set_title("(a)  Overall appraisal by MMAT category", fontsize=9.6, loc="left", pad=8)
    n_und = sum(1 for r in q if catcode(r["mmat_category"]) == "?")
    ax.text(1.0, 1.02, f"n = {len(q)} appraised" +
            (f" · {n_und} undetermined (scanned original)" if n_und else ""),
            transform=ax.transAxes, ha="right", fontsize=7.0, color=T.AXIS)

    # ── (b) 무엇이 잘 보고되고 무엇이 보고되지 않는가 ────────────────
    # ★ 종전 판은 25문항 전건을 실어 메시지가 묻혔다("Figure 8 이해가 잘 안 된다").
    #   본문이 실제로 논하는 문항만 남기고, **잘 보고된 것 / 보고되지 않은 것**을
    #   두 묶음으로 갈라 논지를 그림에서 바로 읽히게 한다. 전건은 Supplementary S10.
    ax = axes[1]
    GOOD = [("3.2", "Non-randomised"), ("3.5", "Non-randomised"),
            ("4.1", "Descriptive"), ("1.1", "Qualitative")]
    POOR = [("4.2", "Descriptive"), ("4.4", "Descriptive"),
            ("3.4", "Non-randomised"), ("3.1", "Non-randomised"),
            ("2.1", "RCT"), ("2.2", "RCT"), ("2.4", "RCT")]
    blocks = [("Reported well — about the measurement", GOOD),
              ("Reported poorly or not at all — about the people", POOR)]

    rows, ylabels, heads = [], [], []
    slot = 0.0
    for head, keys in blocks:
        heads.append((slot, head)); slot += 1.0
        for k, cat in keys:
            c = item_v.get(k)
            if not c:
                continue
            tot = sum(c.values()) or 1
            rows.append((slot, [c.get("Y", 0) / tot, c.get("CT", 0) / tot, c.get("N", 0) / tot],
                         c, tot))
            ylabels.append((slot, f"{k}  {ITEMS[k]}", cat))
            slot += 1.0
        slot += 0.6
    slot -= 0.6            # 마지막 블록 뒤 여백 제거

    ymax = slot
    for ypos, fracs, c, tot in rows:
        yy = ymax - ypos
        left = 0.0
        for frac, col in zip(fracs, (YES, CT, NO)):
            if frac > 0:
                ax.barh(yy, frac, left=left, height=0.62, color=col,
                        edgecolor=T.SURF, linewidth=0.8, zorder=3)
                if frac >= 0.14:
                    ax.text(left + frac / 2, yy, f"{frac*100:.0f}%", ha="center", va="center",
                            fontsize=6.6, color=T.SURF if col != CT else T.INK, zorder=5)
            left += frac
        ax.text(1.012, yy, f"{c.get('Y', 0)}/{tot}", va="center", ha="left",
                fontsize=6.8, color=T.INK2, transform=ax.get_yaxis_transform())

    ax.set_yticks([ymax - p for p, _, _ in ylabels])
    ax.set_yticklabels([lab for _, lab, _ in ylabels], fontsize=7.4)
    for ypos, head in heads:
        ax.text(-0.005, ymax - ypos, head, ha="right", va="center", fontsize=7.8,
                fontweight="bold", color=T.INK, transform=ax.get_yaxis_transform())
    ax.set_xlim(0, 1); ax.set_ylim(-0.4, ymax + 0.4)
    ax.set_xticks([0, 0.25, 0.5, 0.75, 1.0])
    ax.set_xticklabels(["0%", "25%", "50%", "75%", "100%"], fontsize=7.4)
    ax.set_xlabel("share of studies in the category that the item applies to", fontsize=8)
    for s in ("top", "right", "left"):
        ax.spines[s].set_visible(False)
    ax.tick_params(axis="y", length=0)
    handles = [plt.Rectangle((0, 0), 1, 1, color=c) for c in (YES, CT, NO)]
    ax.legend(handles, ["Yes — criterion met", "Can't tell — not reported", "No — not met"],
              fontsize=7.2, loc="upper center", bbox_to_anchor=(0.5, 1.10), ncol=3,
              handlelength=1.0, columnspacing=1.4)
    ax.set_title("(b)  Selected MMAT items — full item set in Supplementary S10",
                 fontsize=9.6, loc="left", pad=22)

    # ★ 그림에 제목을 넣지 않는다 — 캡션이 담당한다(저널 관행).
    fig.subplots_adjust(left=0.315, right=0.975, top=0.955, bottom=0.045)
    for ext in ("png", "pdf"):
        fig.savefig(os.path.join(FIG, f"Fig8_Quality.{ext}"), dpi=300)
    plt.close(fig)
    print("[저장] figures/Fig8_Quality.png|pdf")
    print("  범주별 등급:", {dict(CATS)[c]: dict(tier_by_cat[c]) for c in codes})
    worst = sorted(((item_v[k].get("N", 0) + item_v[k].get("CT", 0)) / max(sum(item_v[k].values()), 1), k)
                   for k in item_v)[-5:]
    print("  감점률 상위:", [(k, f"{f*100:.0f}%") for f, k in reversed(worst)])


if __name__ == "__main__":
    main()
