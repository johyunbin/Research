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
    fig, axes = plt.subplots(2, 1, figsize=(8.4, 8.4),
                             gridspec_kw={"height_ratios": [1.25, 1.0], "hspace": 0.34})

    # (a) 국가 — 'Not reported'는 순위에서 빼고 캡션으로 보고
    ax = axes[0]
    TOPN = 12
    n_nr = cc.get("Not reported", 0)
    items = [(k, v) for k, v in cc.most_common() if k != "Not reported"]
    top = items[:TOPN]
    rest = items[TOPN:]
    labels = [k for k, _ in top][::-1]
    vals = [v for _, v in top][::-1]
    if rest:
        labels = [f"Other ({len(rest)} countries)"] + labels
        vals = [sum(v for _, v in rest)] + vals
    y = np.arange(len(vals))
    ax.barh(y, vals, height=0.62, color=T.BLUE, edgecolor="none", zorder=3)
    ax.set_yticks(y); ax.set_yticklabels(labels, fontsize=8)
    ax.set_xlabel("studies (n)", fontsize=8)
    ax.set_xlim(0, max(vals) * 1.16)
    ax.grid(axis="x", zorder=0); ax.set_axisbelow(True)
    for s in ("top", "right", "left"):
        ax.spines[s].set_visible(False)
    tot = len(rows)
    for yi, v in zip(y, vals):
        ax.text(v + max(vals) * 0.012, yi, str(v), va="center", ha="left",
                fontsize=7.4, color=T.INK2)
    top_country, top_n = items[0]
    ax.annotate(f"{top_country}: {top_n} of {tot} studies ({top_n/tot*100:.0f}%)",
                xy=(top_n, len(vals) - 1), xytext=(top_n * 0.52, len(vals) - 3.5),
                fontsize=8.0, color=T.ORANGE, fontweight="bold",
                arrowprops=dict(arrowstyle="-|>", color=T.ORANGE, lw=1.3,
                                connectionstyle="arc3,rad=0.25"), zorder=6)
    ax.set_title("(a)  Where the evidence comes from", fontsize=9.6, loc="left", pad=8)
    ax.text(1.0, 1.02, f"multi-country studies counted once per country · "
            f"{n_nr} did not report a country",
            transform=ax.transAxes, ha="right", fontsize=7.0, color=T.AXIS)

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

    cats = [("forward", T.BLUE, "forward — sound → behaviour"),
            ("reverse", T.ORANGE, "reverse — behaviour → sound"),
            ("both", T.MUTED, "both directions")]
    x = np.arange(len(span), dtype=float)
    x[1:] += 0.55                     # 압축 칸과 연도축 사이 시각적 분리
    bottom = np.zeros(len(span))
    for key, col, lab in cats:
        v = np.array([cnt(s, key) for s in span], float)
        ax.bar(x, v, bottom=bottom, width=0.74, color=col, edgecolor=T.SURF,
               linewidth=0.6, label=lab, zorder=3)
        bottom += v
    ax.axvline(0.78, color=T.GRID, lw=0.9, zorder=1)
    ax.set_ylabel("studies (n)", fontsize=8)
    ax.grid(axis="y", zorder=0); ax.set_axisbelow(True)
    for s in ("top", "right"):
        ax.spines[s].set_visible(False)
    ax.set_xticks(x)
    ax.set_xticklabels(["≤2009"] + [(str(t) if t % 2 == 0 else "")
                                    for t in range(CUT, y1 + 1)], fontsize=7.4)
    ax.set_xlim(-0.8, x[-1] + 0.8)
    n_pre = sum(sum(yr[y].values()) for y in years if y < CUT)
    ax.text(0, bottom[0] + 0.35, f"{n_pre} studies\n1978–2009", ha="center", va="bottom",
            fontsize=6.8, color=T.INK2, linespacing=1.4)
    ax.legend(loc="upper left", fontsize=7.4, ncol=1, handlelength=1.1,
              borderpad=0.2, labelspacing=0.35)
    ax.set_title("(b)  When it was published, and which direction it tested",
                 fontsize=9.6, loc="left", pad=8)
    recent = sum(sum(yr[y].values()) for y in years if y >= 2020)
    ax.text(1.0, 1.02, f"{recent} of {tot} studies ({recent/tot*100:.0f}%) published since 2020",
            transform=ax.transAxes, ha="right", fontsize=7.0, color=T.AXIS)

    fig.suptitle("Geographic and temporal distribution of the included studies",
                 fontsize=11.4, x=0.012, ha="left", y=0.985)
    fig.text(0.012, 0.955, f"n = {tot} studies (98 included + 4 reserved for sensitivity). "
             "Panel (a) is the basis for the generalisability caveat in the Discussion.",
             fontsize=7.8, color=T.INK2, ha="left")
    fig.subplots_adjust(left=0.20, right=0.975, top=0.905, bottom=0.062)
    for ext in ("png", "pdf"):
        fig.savefig(os.path.join(FIG, f"Fig7_GeoTime.{ext}"), dpi=300)
    plt.close(fig)

    print(f"[저장] figures/Fig7_GeoTime.png|pdf · fulltext/geo_time_counts.csv")
    print(f"  국가 {len(cc)}종 · 상위: {items[:6]}")
    print(f"  Not reported {cc.get('Not reported', 0)}건")
    if unmapped:
        print(f"  ⚠️ 국가 파싱 실패 원본값: {dict(unmapped)}")
    print(f"  연도 {min(years)}~{y1} · 2009년 이전 {n_pre}편 압축 · 2020년 이후 {recent}편")


if __name__ == "__main__":
    main()
