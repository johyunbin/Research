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

    # 2026-09-17 사용자: "(b) 의 상하 폭을 좀더 늘리자. 약간 납작한 느낌" → 전체 5.4→7.0 in, (b) 비중 1.62→1.95
    fig, axes = plt.subplots(2, 1, figsize=(T.W_FULL, 7.0),
                             gridspec_kw={"height_ratios": [1.0, 1.95], "hspace": 0.42})

    # ── (a) 범주별 등급 ─────────────────────────────────────────────
    ax = axes[0]
    codes = [c for c, _ in CATS if tier_by_cat.get(c)]
    labels = [f"{dict(CATS)[c]}" for c in codes]
    # 등급은 순서형 — 단일 hue 3단 램프(어두움 = high)
    tiers = [("high", T.ORD3[2], "high (4–5 criteria met)"),
             ("moderate", T.ORD3[1], "moderate (3)"),
             ("low", T.ORD3[0], "low (0–2)")]
    y = np.arange(len(codes))[::-1].astype(float)
    left = np.zeros(len(codes))
    for key, col, lab in tiers:
        v = np.array([tier_by_cat[c].get(key, 0) for c in codes], float)
        ax.barh(y, v, left=left, height=0.6, color=col, edgecolor=T.SURF,
                linewidth=0.7, label=lab, zorder=3)
        for yi, (l0, vi) in enumerate(zip(left, v)):
            if vi >= 2:
                ax.text(l0 + vi / 2, y[yi], f"{int(vi)}", ha="center", va="center",
                        fontsize=7.6, color=T.ink_on(col), fontweight="bold", zorder=5)
        left += v
    totals = left
    for yi, tot in zip(y, totals):
        ax.text(tot + 0.7, yi, f"n = {int(tot)}", va="center", fontsize=7.5, color=T.INK2)
    ax.set_yticks(y); ax.set_yticklabels(labels, fontsize=8)
    ax.set_xlim(0, max(totals) * 1.18)
    ax.xaxis.set_visible(False)               # 세그먼트 값이 전부 인쇄돼 있다
    for s in ("top", "right", "left", "bottom"):
        ax.spines[s].set_visible(False)
    ax.tick_params(axis="y", length=0)
    ax.legend(fontsize=7.5, loc="upper right", ncol=1, handlelength=1.0,
              columnspacing=1.1, borderpad=0.2, labelspacing=0.34)
    ax.set_title("(a)  Overall appraisal by MMAT category", fontsize=9, loc="left",
                 pad=8, fontweight="bold")
    n_und = sum(1 for r in q if catcode(r["mmat_category"]) == "?")
    # 주석은 패널 오른쪽 아래 빈 공간에 — 제목 줄에 두면 겹친다(실측)
    ax.text(1.0, 0.02, f"n = {len(q)} appraised" +
            (f" · {n_und} not categorised" if n_und else ""),
            transform=ax.transAxes, ha="right", va="bottom", fontsize=7.2, color=T.AXIS)

    # ── (b) 무엇이 잘 보고되고 무엇이 보고되지 않는가 ────────────────
    # ★ 종전 판은 25문항 전건을 실어 메시지가 묻혔다("Figure 8 이해가 잘 안 된다").
    #   본문이 실제로 논하는 문항만 남기고, **잘 보고된 것 / 보고되지 않은 것**을
    #   두 묶음으로 갈라 논지를 그림에서 바로 읽히게 한다. 전건은 OSF 품질평가 자료.
    #   2026-09-17: 묶음 이름을 해석("Reported well/poorly")에서 사실(충족 비율 구간)로 바꿨다 —
    #   구간이 데이터와 어긋나면 아래 검사가 멈춘다.
    ax = axes[1]
    GOOD = [("3.2", "Non-randomised"), ("3.5", "Non-randomised"),
            ("4.1", "Descriptive"), ("1.1", "Qualitative")]
    POOR = [("4.2", "Descriptive"), ("4.4", "Descriptive"),
            ("3.4", "Non-randomised"), ("3.1", "Non-randomised"),
            ("2.1", "RCT"), ("2.2", "RCT"), ("2.4", "RCT")]
    blocks = [("Met in ≥80% of studies", GOOD),
              ("Met in <50% of studies", POOR)]
    for keys, ok in ((GOOD, lambda s: s >= 0.8), (POOR, lambda s: s < 0.5)):
        for k, _ in keys:
            c = item_v.get(k) or {}
            share = c.get("Y", 0) / (sum(c.values()) or 1)
            if not ok(share):
                raise SystemExit(f"⚠️ MMAT 문항 {k} 충족 비율 {share:.0%} 가 묶음 이름과 맞지 않는다")

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
                ax.barh(yy, frac, left=left, height=0.64, color=col,
                        edgecolor=T.SURF, linewidth=0.7, zorder=3)
                if frac >= 0.14:
                    ax.text(left + frac / 2, yy, f"{frac*100:.0f}%", ha="center", va="center",
                            fontsize=7.2, color=T.ink_on(col), zorder=5)
            left += frac
        ax.text(1.014, yy, f"{c.get('Y', 0)}/{tot}", va="center", ha="left",
                fontsize=7.2, color=T.INK2, transform=ax.get_yaxis_transform())

    ax.set_yticks([ymax - p for p, _, _ in ylabels])
    ax.set_yticklabels([lab for _, lab, _ in ylabels], fontsize=7.6)
    for ypos, head in heads:
        ax.text(-0.005, ymax - ypos, head, ha="right", va="center", fontsize=8,
                fontweight="bold", color=T.INK, transform=ax.get_yaxis_transform())
    ax.set_xlim(0, 1); ax.set_ylim(-0.4, ymax + 0.4)
    ax.set_xticks([0, 0.25, 0.5, 0.75, 1.0])
    ax.set_xticklabels(["0%", "25%", "50%", "75%", "100%"], fontsize=7.5)
    ax.set_xlabel("share of studies in the category that the item applies to", fontsize=8)
    for s in ("top", "right", "left"):
        ax.spines[s].set_visible(False)
    ax.tick_params(axis="y", length=0)
    handles = [plt.Rectangle((0, 0), 1, 1, color=c) for c in (YES, CT, NO)]
    leg_b = ax.legend(handles, ["Yes", "Can't tell", "No"],
                      fontsize=7.5, loc="upper center", bbox_to_anchor=(0.5, 1.115), ncol=3,
                      handlelength=1.0, columnspacing=1.3)
    ax.set_title("(b)  Selected MMAT items",
                 fontsize=9, loc="left", pad=24, fontweight="bold")

    # ★ 그림에 제목을 넣지 않는다 — 캡션이 담당한다(저널 관행).
    fig.subplots_adjust(left=0.345, right=0.955, top=0.955, bottom=0.05)
    # 범례는 패널 위 0.27 in 에 고정 — 축 좌표(1.115)로 두면 패널을 키울수록 제목 쪽으로 올라간다
    h_in = ax.get_position().height * fig.get_figheight()
    leg_b.set_bbox_to_anchor((0.5, 1 + 0.27 / h_in))
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
