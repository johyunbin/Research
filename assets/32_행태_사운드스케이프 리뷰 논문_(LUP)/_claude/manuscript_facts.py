# -*- coding: utf-8 -*-
"""
Paper32 — 원고에 들어갈 수치를 정본 데이터에서 전부 재산출
원고를 96편 → 98편 기준으로 고칠 때 손으로 세지 않기 위한 단일 출처.
출력: fulltext/manuscript_facts.md
"""
import sys, os, csv, re, statistics
from collections import Counter, defaultdict

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")


def rd(p):
    return list(csv.DictReader(open(os.path.join(FT, p), encoding="utf-8-sig")))


def main():
    ext = rd("corpus_v4_extraction.csv")
    verd = {r["uid"]: r["final_verdict"] for r in rd("corpus_v4_verdicts.csv")}
    qual = {r["uid"]: r for r in rd("quality_v2.csv")}
    t1 = {r["uid"]: r for r in rd("table1_v2.csv")}
    inc = [r for r in ext if verd.get(r["uid"]) == "FINAL_INCLUDE"]
    N = len(inc)
    L = [f"# 원고 수치 정본 (포함 {N}편 · 재현 = `manuscript_facts.py`)\n\n"]

    def add(k, v, note=""):
        L.append(f"- **{k}** = {v}" + (f"  ({note})\n" if note else "\n"))
        print(f"{k:44s} {v}")

    # ── 연도 ──
    yrs = sorted(int(r["year"]) for r in inc if (r["year"] or "").strip().isdigit())
    y2020 = sum(1 for y in yrs if y >= 2020)
    L.append("\n## 1. 시기\n\n")
    add("포함 편수", N)
    add("2020년 이후", f"{y2020} / {len(yrs)} ({y2020/len(yrs)*100:.0f}%)")
    add("중앙 출판연도", int(statistics.median(yrs)))
    add("최고연도", min(yrs))

    # ── 표본 ────────────────────────────────────────────────────────
    # ⚠️ 단순히 첫 숫자를 뽑으면 "Study 1 = 105명…" 이 1이 되고 "NR (Table 1 …)" 이 1이 된다.
    #    ① NR 시작 = 미보고로 제외 ② 관측 단위가 사람이 아닌 연구(가로구간·GPS궤적·SVI)는 분리.
    NONPART = r"segment|point|trip|route|image|svi|trajector|sample point|가로|궤적|표본점"
    part, nonpart, nr = [], [], []
    for r in inc:
        s = (r["sample_n"] or "").strip()
        if not s or re.match(r"^\s*(NR|N/?A|-)\b", s, re.I):
            nr.append(r["uid"]); continue
        m = re.match(r"^\s*([\d,]+)", s)          # 문두 숫자만 신뢰한다
        if not m:
            nr.append(r["uid"]); continue
        try: n = int(m.group(1).replace(",", ""))
        except ValueError:
            nr.append(r["uid"]); continue
        (nonpart if re.search(NONPART, s, re.I) else part).append((n, r["uid"], s[:50]))
    ns = sorted(n for n, _, _ in part)
    L.append("\n## 2. 표본\n\n")
    add("참가자 단위 보고", f"{len(part)}편", f"비참가자 단위 {len(nonpart)}편 · 미보고/파싱불가 {len(nr)}편")
    add("참가자 수 중앙값", int(statistics.median(ns)))
    q1 = ns[len(ns) // 4]; q3 = ns[3 * len(ns) // 4]
    add("IQR", f"{q1} – {q3}")
    add("참가자 최대", f"{max(ns):,}")
    nonpart.sort(reverse=True)
    add("최대 비참가자 데이터셋", " / ".join(f"{n:,} ({u})" for n, u, _ in nonpart[:3]))

    # ── 세팅·국가·설계 ──
    def bucket(v, pats):
        s = (v or "").lower()
        return [name for name, p in pats if re.search(p, s)]

    SET = [("street", r"street|road|sidewalk|pavement|거리"), ("park", r"park|garden|green space"),
           ("square", r"square|plaza|piazza"), ("campus", r"campus|university|school"),
           ("residential", r"residential|neighbourhood|neighborhood|housing"),
           ("waterfront", r"waterfront|river|lake|coast|harbour|harbor")]
    setc = Counter()
    for r in inc:
        for b in bucket(r["setting"], SET):
            setc[b] += 1

    # 국가 정규화기는 Fig 7과 **같은 것**을 쓴다 — 따로 쓰면 본문과 그림이 어긋난다(실제로 어긋났다).
    sys.path.insert(0, BASE)
    from make_fig7_geo_time import norm_countries
    ctry = Counter()
    for r in inc:
        cs = norm_countries(r["country"])
        if not cs:
            ctry["Not reported"] += 1
        for c in cs:
            ctry[c] += 1

    # design 필드는 `유형 (부연)` 형태가 섞여 있다 — 괄호 앞 머리말만 본다.
    DES = [("field-experiment", r"^field[\s-]?experiment|^natural[\s-]?experiment"),
           ("lab/VR-experiment", r"^lab|^vr\b|virtual"),
           ("mixed obs+survey", r"^mixed"),
           ("survey", r"^survey|^questionnaire"),
           ("field-observation", r"^field[\s-]?observation|^observation"),
           ("sensor/bigdata", r"^sensor|^big[\s-]?data")]
    des = Counter()
    for r in inc:
        head = (r["design"] or "").lower().split("(")[0].strip()
        for name, p in DES:
            if re.search(p, head):
                des[name] += 1; break
        else:
            des["other/NR"] += 1

    L.append("\n## 3. 세팅·지리·설계\n\n")
    add("세팅", " · ".join(f"{k} {v}" for k, v in setc.most_common()))
    top = ctry.most_common(6)
    add("국가 상위", " · ".join(f"{k} {v}" for k, v in top))
    if top:
        add("최다국 비율", f"{top[0][1]}/{N} ({top[0][1]/N*100:.0f}%)")
    add("설계", " · ".join(f"{k} {v}" for k, v in des.most_common()))
    fe = des["field-experiment"]
    add("현장·자연실험", f"{fe} ({fe/N*100:.0f}%)")
    add("실험 계열 전체", f"{fe+des['lab/VR-experiment']} "
                          f"({(fe+des['lab/VR-experiment'])/N*100:.0f}%)")

    # ── 품질 ──
    qt = Counter(qual[r["uid"]]["quality_tier"] for r in inc if r["uid"] in qual)
    L.append("\n## 4. MMAT 품질(포함분만)\n\n")
    add("tier", f"high {qt['high']} · moderate {qt['moderate']} · low {qt['low']}")

    # MMAT 범주: 1.x 질적 · 2.x RCT · 3.x 비무작위 정량 · 4.x 정량기술 · 5.x 혼합
    CATNAME = {"1": "qualitative", "2": "RCT", "3": "non-randomised",
               "4": "quantitative descriptive", "5": "mixed methods"}
    ITEM = {"3.4": "교란 통제", "4.2": "표본 대표성", "4.4": "무응답 편의 낮음",
            "4.1": "표집전략 적절", "4.3": "측정 적절", "4.5": "통계분석 적절",
            "3.1": "참가자 대표성", "3.3": "완전한 아웃컴 자료"}
    det = [r for r in rd("quality_detail_v2.csv") if verd.get(r["uid"]) == "FINAL_INCLUDE"]
    byq = defaultdict(Counter)
    for r in det:
        if r["item_no"]:
            byq[r["item_no"]][r["verdict"]] += 1
    L.append("\n범주별 편수: " + " · ".join(
        f"{CATNAME[c]} {len({r['uid'] for r in det if r['item_no'].startswith(c + '.')})}"
        for c in "12345") + "\n\n주요 항목(yes / 분모 · CT=보고 없음):\n\n")
    for item in sorted(byq):
        c = byq[item]; tot = sum(c.values())
        line = (f"  - **{item}** {ITEM.get(item, '')} → Y {c['Y']}/{tot} "
                f"({c['Y']/tot*100:.0f}%) · N {c['N']} · CT {c['CT']}")
        L.append(line + "\n"); print(line)

    # ── 방향 ──
    dirc = Counter()
    for r in inc:
        for d in (r["behaviour_domain"] or "").split(";"):
            if d.strip():
                dirc[(r["direction"] or "").strip()] += 1
    tf, tr, tb = dirc["forward"], dirc["reverse"], dirc["both"]
    L.append("\n## 5. 방향\n\n")
    add("도메인 레코드", f"forward {tf} · reverse {tr} · both {tb} (합 {tf+tr+tb})")
    add("역방향 비율(both 제외)", f"{tr/(tf+tr)*100:.0f}%")
    add("역방향 비율(both 포함 분모)", f"{tr/(tf+tr+tb)*100:.0f}%")

    dom = defaultdict(Counter)
    for r in inc:
        for d in (r["behaviour_domain"] or "").split(";"):
            d = d.strip().lower()
            if d:
                dom[d][(r["direction"] or "").strip()] += 1
    for d in ("movement", "staying", "space-use", "social", "activity"):
        add(f"  {d}", f"fwd {dom[d]['forward']} · rev {dom[d]['reverse']} · both {dom[d]['both']}")

    # ── 측정세대 ──
    band = defaultdict(Counter); tot_g = Counter(); multi = 0
    for r in inc:
        gs = {g.strip() for g in (t1.get(r["uid"], {}).get("measure_gen", "") or "").split(";")
              if g.strip() in ("G1", "G2", "G3")}
        if len(gs) >= 2:
            multi += 1
        try: y = int(r["year"])
        except (TypeError, ValueError): continue
        b = "≤2009" if y < 2010 else ("2010-2019" if y < 2020 else "2020-")
        for g in gs:
            band[b][g] += 1; tot_g[g] += 1
    L.append("\n## 6. 측정 세대\n\n")
    for b in ("≤2009", "2010-2019", "2020-"):
        add(f"  {b}", f"G1 {band[b]['G1']} · G2 {band[b]['G2']} · G3 {band[b]['G3']}")
    add("합계", f"G1 {tot_g['G1']} · G2 {tot_g['G2']} · G3 {tot_g['G3']}")
    add("2세대 이상 병용", multi)

    # ── 음원 ──
    air = sum(1 for r in inc if re.search(r"aircraft|airport|항공|flight",
                                          f"{r['exposure']} {r['title']}", re.I))
    L.append("\n## 7. 음원 공백\n\n")
    add("항공기소음 레코드", air)

    # ── 갈래 ──
    src = Counter(r["source"] for r in inc)
    L.append("\n## 8. 갈래\n\n")
    add("갈래별 포함", " · ".join(f"{k} {v}" for k, v in src.items()))
    add("인용추적 기여율", f"{src['citation-tracking']}/{src['db-search']} = "
                          f"{src['citation-tracking']/src['db-search']*100:.0f}%")
    sens = sum(1 for v in verd.values() if v == "SENS_ONLY")
    add("민감도 전용", sens)
    add("분석 총계", N + sens, "FINAL_EXCLUDE는 분모 아님")

    open(os.path.join(FT, "manuscript_facts.md"), "w", encoding="utf-8").write("".join(L))
    print("\n[저장] fulltext/manuscript_facts.md")


if __name__ == "__main__":
    main()
