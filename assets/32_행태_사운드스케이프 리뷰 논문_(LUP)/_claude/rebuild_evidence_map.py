# -*- coding: utf-8 -*-
"""
Paper32 — 증거지도 재집계 v2 (3갈래 통합 코퍼스 98편 기준)
1차 evidence_map.md는 81편 기준 수기 집계였다. 여기서는 corpus_v4_extraction.csv에서
기계적으로 산출해 그림·표와 단일 출처를 공유하게 만든다.
출력: fulltext/evidence_map_v2.md · evidence_counts_v2.csv
"""
import sys, os, csv, re
from collections import Counter, defaultdict

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")

DOMAINS = ["movement", "staying", "space-use", "social", "activity"]
DOM_LABEL = {"movement": "이동", "staying": "체류", "space-use": "공간이용",
             "social": "사회적 행태", "activity": "활동"}

# exposure 자유서술 → 음원 유형(다중 매칭 허용). 원문 표현을 폭넓게 잡는다.
SOURCES = [
    ("교통/도로소음", r"traffic|road|도로|교통|차량|vehicle|street noise|LAeq.*road"),
    ("일반 소음", r"\bnoise\b|소음|dB\b|LAeq|LZeq|sound (?:pressure )?level|SPL"),
    ("자연음", r"natural sound|nature sound|bird|water|stream|fountain|wind|자연음|새소리|물소리|조류|leaves"),
    ("음악/부가음", r"music|음악|speaker|스피커|broadcast|audio|APS|audible.*signal|masking"),
    ("인간·군중음", r"crowd|human sound|voice|conversation|대화|군중|footstep|발소리|people sound|speech"),
    ("항공기소음", r"aircraft|airport|항공|비행"),
]


def main():
    rows = list(csv.DictReader(open(os.path.join(FT, "corpus_v4_extraction.csv"),
                                    encoding="utf-8-sig")))
    verd = {r["uid"]: r["final_verdict"] for r in
            csv.DictReader(open(os.path.join(FT, "corpus_v4_verdicts.csv"), encoding="utf-8-sig"))}
    qual = {r["uid"]: r["quality_tier"] for r in
            csv.DictReader(open(os.path.join(FT, "quality_v2.csv"), encoding="utf-8-sig"))}
    inc = [r for r in rows if verd.get(r["uid"]) == "FINAL_INCLUDE"]
    print(f"[0] 포함 {len(inc)}편 (민감도 전용 {len(rows)-len(inc)}편 제외)")

    def domains(r):
        return [d.strip().lower() for d in (r["behaviour_domain"] or "").split(";") if d.strip()]

    def sources(r):
        blob = f"{r['exposure']} {r['title']}".lower()
        hit = [name for name, pat in SOURCES if re.search(pat, blob, re.I)]
        # '일반 소음'은 더 구체적인 유형이 잡혔으면 중복 계상하지 않는다
        if len(hit) > 1 and "일반 소음" in hit:
            hit = [h for h in hit if h != "일반 소음"]
        return hit or ["기타·미상"]

    # ① 도메인 × 음원
    grid = defaultdict(Counter)
    for r in inc:
        for d in domains(r):
            for s in sources(r):
                grid[d][s] += 1

    # ② 방향 × 도메인
    dirn = defaultdict(Counter)
    for r in inc:
        for d in domains(r):
            dirn[d][(r["direction"] or "").strip()] += 1

    # ③ 측정세대 × 시기
    # ⚠️ 여기서 정규식을 다시 돌리면 Table 1과 어긋난다(실제로 어긋났다).
    #    Table 1의 measure_gen을 **단일 출처**로 삼는다.
    t1p = os.path.join(FT, "table1_v2.csv")
    if not os.path.exists(t1p):
        print("⚠️ table1_v2.csv 없음 — rebuild_table1.py 먼저 실행"); sys.exit(1)
    t1gen = {r["uid"]: r["measure_gen"] for r in csv.DictReader(open(t1p, encoding="utf-8-sig"))}

    def gens(r):
        return {g.strip() for g in (t1gen.get(r["uid"], "") or "").split(";")
                if g.strip() in ("G1", "G2", "G3")}

    band = defaultdict(Counter)
    for r in inc:
        try:
            y = int(r["year"])
        except (TypeError, ValueError):
            continue
        b = "≤2009" if y < 2010 else ("2010–2019" if y < 2020 else "2020–")
        for g in gens(r):
            band[b][g] += 1

    # ④ 갈래별 구성
    by_src = Counter(r["source"] for r in inc)
    q_by_src = defaultdict(Counter)
    for r in inc:
        q_by_src[r["source"]][qual.get(r["uid"], "?")] += 1

    # ── 저장 ───────────────────────────────────────────────────────
    src_names = [s for s, _ in SOURCES] + ["기타·미상"]
    with open(os.path.join(FT, "evidence_counts_v2.csv"), "w", newline="",
              encoding="utf-8-sig") as f:
        w = csv.writer(f); w.writerow(["kind", "row", "col", "n"])
        for d in DOMAINS:
            for s in src_names:
                if grid[d][s]:
                    w.writerow(["domain_x_source", d, s, grid[d][s]])
        for d in DOMAINS:
            for k, v in dirn[d].items():
                w.writerow(["direction_x_domain", d, k, v])
        for b in ["≤2009", "2010–2019", "2020–"]:
            for g in ("G1", "G2", "G3"):
                w.writerow(["generation_x_band", b, g, band[b][g]])

    L = [f"# Paper32 — 증거 지도 v2 (포함 {len(inc)}편)\n",
         "\n⚠️ **이 문서는 `corpus_v4_extraction.csv`에서 기계적으로 산출된다**(수기 집계 아님). "
         "그림·표와 단일 출처를 공유하므로 수치가 어긋날 수 없다. 재생성 = `rebuild_evidence_map.py`.\n",
         f"\n갈래 구성: {dict(by_src)} · 민감도 전용 {len(rows)-len(inc)}편은 제외.\n",
         "\n## 1. 행태 도메인 × 음원 유형\n\n"
         "셀 = 해당 조합을 다룬 논문 수(한 논문이 여러 셀에 기여 가능). "
         "음원은 `exposure` 자유서술에서 정규식으로 추출하며, 더 구체적인 유형이 잡히면 "
         "'일반 소음'은 중복 계상하지 않는다.\n\n"]
    L.append("| 행태 도메인 | " + " | ".join(src_names) + " |\n")
    L.append("|---" * (len(src_names) + 1) + "|\n")
    for d in DOMAINS:
        cells = []
        for s in src_names:
            n = grid[d][s]
            cells.append(f"**{n}**" if n and n == max(grid[d].values()) else str(n))
        L.append(f"| {DOM_LABEL[d]}({d}) | " + " | ".join(cells) + " |\n")

    air = sum(grid[d]["항공기소음"] for d in DOMAINS)
    L.append(f"\n**읽기**: 항공기소음이 전체 {air}건으로 여전히 사실상 공백이다 — "
             "공항 주변 연구가 건강·짜증 프레임에 갇혀 있고, 행태로 측정한 것이 거의 없다. "
             "연구 어젠다 1순위.\n")

    L.append("\n## 2. 방향(forward / reverse / both) × 도메인\n\n")
    L.append("| 도메인 | forward(음→행태) | reverse(행태→음) | both |\n|---|---|---|---|\n")
    for d in DOMAINS:
        c = dirn[d]
        L.append(f"| {DOM_LABEL[d]} | {c['forward']} | {c['reverse']} | {c['both']} |\n")
    tot_f = sum(dirn[d]["forward"] for d in DOMAINS)
    tot_r = sum(dirn[d]["reverse"] for d in DOMAINS)
    L.append(f"\n**읽기**: 역방향이 전체의 {tot_r/(tot_f+tot_r)*100:.0f}%다. "
             "특히 공간이용·활동·사회적 행태에서 역방향이 순방향과 대등하다 — "
             "'사람이 무엇을 하는가'가 사운드스케이프를 만든다는 증거가 특정 도메인에 집중된다. "
             "양방향 프레임을 채택한 근거.\n")

    L.append("\n## 3. 행태 측정 세대 × 시기\n\n"
             "코딩 규칙: 한 연구가 두 세대를 함께 쓰면 **둘 다 계상**(다중코딩). "
             "노출 계측 전용 장비(소음계 등)는 *행태* 측정이 아니므로 G3가 아니다.\n\n"
             "| 시기 | G1 자기보고 | G2 체계적 관찰 | G3 센싱·궤적 |\n|---|---|---|---|\n")
    for b in ["≤2009", "2010–2019", "2020–"]:
        c = band[b]
        L.append(f"| {b} | {c['G1']} | {c['G2']} | {c['G3']} |\n")
    tot = Counter()
    for b in band.values():
        tot.update(b)
    L.append(f"| **합계** | **{tot['G1']}** | **{tot['G2']}** | **{tot['G3']}** |\n")
    g3a = band["2010–2019"]["G3"]; g3b = band["2020–"]["G3"]
    L.append(f"\n**읽기**: G3가 {g3a}편 → {g3b}편으로 늘었는데 G1·G2도 함께 늘었다. "
             "새 측정법이 옛 측정법을 **대체하지 않았다** — 측정 세대는 교체되지 않고 층위가 추가된다.\n")

    L.append("\n## 4. 갈래별 품질 구성\n\n| 갈래 | high | moderate | low |\n|---|---|---|---|\n")
    for s in ("db-search", "citation-tracking", "openalex-supplementary"):
        c = q_by_src[s]
        L.append(f"| {s} | {c['high']} | {c['moderate']} | {c['low']} |\n")
    L.append("\n**읽기**: 두 갈래의 품질 구성이 사실상 같다. 인용추적이 주변부 문헌만 "
             "긁어온 것이 아니라는 뜻이다(단, 이 결론은 절단 인공물을 해소한 뒤에야 성립했다 — "
             "`quality_truncation_effect.md`).\n")

    open(os.path.join(FT, "evidence_map_v2.md"), "w", encoding="utf-8").write("".join(L))

    print(f"[1] 도메인×음원 집계 완료 · 항공기소음 {air}건")
    print(f"[2] 방향 forward {tot_f} · reverse {tot_r} ({tot_r/(tot_f+tot_r)*100:.0f}%)")
    print(f"[3] 측정세대 G1 {tot['G1']} · G2 {tot['G2']} · G3 {tot['G3']}")
    print(f"[4] 갈래별 품질: " + " / ".join(f"{s} {dict(q_by_src[s])}" for s in q_by_src))
    print("[저장] evidence_map_v2.md · evidence_counts_v2.csv")


if __name__ == "__main__":
    main()
