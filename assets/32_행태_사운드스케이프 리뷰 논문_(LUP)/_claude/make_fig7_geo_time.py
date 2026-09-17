# -*- coding: utf-8 -*-
"""
Paper32 — Figure 7: 지리·시기 분포 (Limitations 절 근거)
(a) 국가별 연구 수 — China 편중이 일반화 제약의 핵심
(b) 연도별 게재 수 × 방향(forward / reverse / both)
정사각 · viz_theme 검증 팔레트 · 도면 텍스트 영문.
출력: figures/Fig7_GeoTime.png|pdf · fulltext/geo_time_counts.csv
"""
import sys, os, csv, re
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

# 자유서술 국가 필드 → 표준 국가명. 매핑에 없으면 NR 처리(응답에 보고).
NORM = {
    "uk": "United Kingdom", "united kingdom": "United Kingdom", "england": "United Kingdom",
    "usa": "United States", "united states": "United States",
    "south korea": "South Korea", "korea": "South Korea",
    "czech republic": "Czechia", "czechia": "Czechia",
    "turkiye": "Türkiye", "turkey": "Türkiye",
}
DROP = re.compile(r"(저자\s*소속|참가자|조사도시|실험실|기준|미명시|based in|run online|"
                  r"focus groups|anechoic|measurements in|NR)", re.I)


def norm_countries(raw):
    # ★ 수집국 미보고(NR…)는 괄호 속 저자 소속 국가를 건지지 않는다 — 소속국은 연구 수행국이 아니다
    #   (2026-09-17 독립 검토 적발: CT0025 → United States, 1149 → Malaysia 로 잘못 집계되고 있었다)
    if (raw or "").strip().upper().startswith("NR"):
        return []
    out = []
    for p in re.split(r"[;,/]| and ", raw or ""):
        p = re.sub(r"\(.*?\)", " ", p)
        p = re.sub(r"[^\w\sÀ-ÿ]", " ", p)
        p = re.sub(r"\s+", " ", p).strip()
        if not p:
            continue
        if DROP.search(p):
            # 괄호 안 부가설명에서 실제 국가명만 건지기
            m = re.search(r"\b(Japan|Malaysia|China|UK|United Kingdom|Turkiye|Switzerland)\b", p, re.I)
            if not m:
                continue
            p = m.group(1)
        key = p.lower()
        out.append(NORM.get(key, p.title() if p.islower() else p))
    return sorted(set(out))


def main():
    rows = list(csv.DictReader(open(os.path.join(FT, "corpus_v4_extraction.csv"), encoding="utf-8-sig")))
    # ⚠️ 민감도 전용 4편을 섞으면 그림과 본문이 다른 분모를 쓴다(독립 게이트 지적).
    #    본문 §3.2와 같은 FINAL_INCLUDE 98편으로 맞춘다.
    verd = {r["uid"]: r["final_verdict"] for r in
            csv.DictReader(open(os.path.join(FT, "corpus_v4_verdicts.csv"), encoding="utf-8-sig"))}
    rows = [r for r in rows if verd.get(r["uid"]) == "FINAL_INCLUDE"]
    from collections import Counter, defaultdict
    cc, unmapped = Counter(), Counter()
    for r in rows:
        cs = norm_countries(r["country"])
        if not cs:
            cc["Not reported"] += 1
            unmapped[r["country"][:40]] += 1
        for c in cs:
            cc[c] += 1

    yr = defaultdict(lambda: Counter())
    for r in rows:
        try:
            y = int(r["year"])
        except (TypeError, ValueError):
            continue
        yr[y][(r["direction"] or "forward").strip()] += 1

    with open(os.path.join(FT, "geo_time_counts.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f); w.writerow(["kind", "key", "sub", "n"])
        for k, v in cc.most_common():
            w.writerow(["country", k, "", v])
        for y in sorted(yr):
            for d, n in yr[y].items():
                w.writerow(["year", y, d, n])

    # ── 그림 ────────────────────────────────────────────────────────
    # ★ v2: 정사각 2단 적층 → **가로 병렬**. (a)는 최대 막대(중국) 하나가 축을 지배해
    #   오른쪽이 텅 비었었다 — 폭을 (a) 40 : (b) 60 으로 나눠 빈 공간을 없앤다.
    #   주황 화살표 주석은 장식이라 뺐다. 중국 막대는 진한 단계로, 라벨은 수치에 병기.
    fig, axes = plt.subplots(1, 2, figsize=(T.W_FULL, 3.0),
                             gridspec_kw={"width_ratios": [1.0, 1.45], "wspace": 0.52})

    # (a) 국가 — 'Not reported'는 순위에서 빼고 캡션으로 보고
    ax = axes[0]
    TOPN = 12          # Other 가 China 와 맞먹지 않게 — 편중이 그림의 논지다
    n_nr = cc.get("Not reported", 0)
    items = [(k, v) for k, v in cc.most_common() if k != "Not reported"]
    top = items[:TOPN]
    rest = items[TOPN:]
    labels = [k for k, _ in top][::-1]
    vals = [v for _, v in top][::-1]
    if rest:
        labels = [f"Other ({len(rest)})"] + labels
        vals = [sum(v for _, v in rest)] + vals
    y = np.arange(len(vals))
    tot = len(rows)
    top_country, top_n = items[0]
    cols = [T.DEEP if lab == top_country else T.BLUE for lab in labels]
    ax.barh(y, vals, height=0.66, color=cols, edgecolor="none", zorder=3)
    ax.set_yticks(y); ax.set_yticklabels(labels, fontsize=8)
    ax.set_xlim(0, max(vals) * 1.30)
    ax.xaxis.set_visible(False)               # 값이 전부 인쇄돼 있어 눈금이 중복이다
    for s in ("top", "right", "left", "bottom"):
        ax.spines[s].set_visible(False)
    ax.tick_params(axis="y", length=0)
    for yi, v, lab in zip(y, vals, labels):
        txt = f"{v}  ({v/tot*100:.0f}%)" if lab == top_country else str(v)
        ax.text(v + max(vals) * 0.03, yi, txt, va="center", ha="left", fontsize=7.5,
                color=(T.DEEP if lab == top_country else T.INK2),
                fontweight=("bold" if lab == top_country else "normal"))
    ax.set_title("(a)  Country", fontsize=9, loc="left", pad=8, fontweight="bold")

    # (b) 연도 × 방향 — 2009년 이전은 한 칸으로 압축(1978~2003 공백 제거)
    ax = axes[1]
    years = sorted(yr)
    y1 = max(years)
    CUT = 2010
    span = ["≤2009"] + [str(t) for t in range(CUT, y1 + 1)]

    def cnt(slot, key):
        if slot == "≤2009":
            return sum(yr[y].get(key, 0) for y in years if y < CUT)
        return yr[int(slot)].get(key, 0)

    cats = [("forward", T.BLUE, "forward"),
            ("both", T.NEUT, "both"),
            ("reverse", T.TERRA, "reverse")]
    x = np.arange(len(span), dtype=float)
    x[1:] += 0.55                     # 압축 칸과 연도축 사이 시각적 분리
    bottom = np.zeros(len(span))
    for key, col, lab in cats:
        v = np.array([cnt(s, key) for s in span], float)
        ax.bar(x, v, bottom=bottom, width=0.76, color=col, edgecolor=T.SURF,
               linewidth=0.5, label=lab, zorder=3)
        bottom += v
    ax.axvline(0.78, color=T.GRID, lw=0.9, zorder=1)
    ax.set_ylabel("studies (n)", fontsize=8)
    ax.grid(axis="y", zorder=0); ax.set_axisbelow(True)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.set_xticks(x)
    ax.set_xticklabels(["≤'09"] + [(f"'{t % 100:02d}" if t % 2 == 0 else "")
                                   for t in range(CUT, y1 + 1)], fontsize=7.5)
    ax.set_xlim(-0.8, x[-1] + 0.8)
    ax.legend(loc="upper left", fontsize=7.5, ncol=1, handlelength=1.0,
              borderpad=0.2, labelspacing=0.35)
    ax.set_title("(b)  Year × direction", fontsize=9, loc="left", pad=8,
                 fontweight="bold")
    recent = sum(sum(yr[y].values()) for y in years if y >= 2020)
    ax.text(1.0, 1.035, f"{recent/tot*100:.0f}% since 2020",
            transform=ax.transAxes, ha="right", va="bottom", fontsize=7.2, color=T.AXIS)

    fig.subplots_adjust(left=0.155, right=0.985, top=0.90, bottom=0.09)
    for ext in ("png", "pdf"):
        fig.savefig(os.path.join(FIG, f"Fig7_GeoTime.{ext}"), dpi=300)
    plt.close(fig)

    print(f"[저장] figures/Fig7_GeoTime.png|pdf · fulltext/geo_time_counts.csv")
    print(f"  국가 {len(cc)}종 · 상위: {items[:6]}")
    print(f"  Not reported {cc.get('Not reported', 0)}건")
    if unmapped:
        print(f"  ⚠️ 국가 파싱 실패 원본값: {dict(unmapped)}")
    n_pre = sum(sum(yr[y].values()) for y in years if y < CUT)
    print(f"  연도 {min(years)}~{y1} · 2009년 이전 {n_pre}편 압축 · 2020년 이후 {recent}편")


if __name__ == "__main__":
    main()
