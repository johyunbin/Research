# -*- coding: utf-8 -*-
"""
Paper32 — 코퍼스 반영(포함 113·민감도 9) 원고 수치·문장 갱신: 메타분석과 무관한 부분 (2026-09-18)

- 수치는 정본 파일(table1_v2.csv·corpus_v4_extraction.csv·quality_v2.csv·evidence_counts_v2.csv·manuscript_facts.md)에서 계산하거나,
  같은 파일에서 확인한 값을 쓴다. 모든 치환은 원문이 정확히 1회 있을 때만 수행(없거나 여러 번이면 중단).
- 서론의 자기 결과 수치 문장은 LUP 리뷰 관행(검토 보고서 B3)에 따라 수치를 고치지 않고 삭제한다.
- 메타분석(3.4·3.5, 초록·결론의 클러스터 문장, Table 2·3)은 별도 단계에서 갱신한다.
"""
import csv, os, sys
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
P = os.path.join(os.path.dirname(BASE), "01_논문작업", "Manuscript_KO.md")
sys.path.insert(0, BASE)
from build_appendix import norm_countries  # noqa: E402

t = open(P, encoding="utf-8").read()


def rep(a, b):
    global t
    n = t.count(a)
    if n != 1:
        sys.exit(f"STOP count={n}: {a[:80]}")
    t = t.replace(a, b)


# ── 정본에서 계산 ─────────────────────────────────────────────
t1 = [r for r in csv.DictReader(open(os.path.join(FT, "table1_v2.csv"), encoding="utf-8-sig")) if r["verdict"] == "FINAL_INCLUDE"]
N = len(t1)
assert N == 113, N
pct = lambda n: f"{round(100 * n / N)}%"
per = Counter("≤2009" if int(r["year"]) <= 2009 else "2010–2019" if int(r["year"]) <= 2019 else "2020–2026" for r in t1)
ext = {r["uid"]: r for r in csv.DictReader(open(os.path.join(FT, "corpus_v4_extraction.csv"), encoding="utf-8-sig"))}
cc = Counter()
for r in t1:
    for x in set(norm_countries(ext[r["uid"]]["country"]) or ["Not reported"]):
        cc[x] += 1
q = {r["uid"]: r for r in csv.DictReader(open(os.path.join(FT, "quality_v2.csv"), encoding="utf-8-sig"))}
tier = Counter(q[r["uid"]]["quality_tier"] for r in t1)
dirc = Counter(r["direction"] for r in t1)
gen = Counter(g for r in t1 for g in r["measure_gen"].split(";") if g in ("G1", "G2", "G3"))
rev = [r for r in t1 if r["direction"] == "reverse"]
# Table 1 세팅·설계 집계는 rebuild_table1/build_tables_en 과 같은 값(table1_en.md)을 쓴다
t1en = open(os.path.join(FT, "table1_en.md"), encoding="utf-8").read()


def cnt(label):
    for line in t1en.splitlines():
        if line.startswith(f"| {label} |"):
            return int(line.split("|")[2])
    sys.exit(f"table1_en.md 에 {label} 없음")


setting = [(k, cnt(k)) for k in ("Street", "Park", "Square / plaza", "Waterfront", "Other", "Campus", "Residential open space")]
design = [(k, cnt(k)) for k in ("Mixed observation + survey", "Survey", "Field or natural experiment", "Field observation",
                                "Sensor / big data", "Laboratory / VR experiment", "Other / not reported")]
assert sum(per.values()) == N and sum(v for _, v in design) == N and sum(tier.values()) == N and sum(dirc.values()) == N
print("period", dict(per), "| tier", dict(tier), "| dir", dict(dirc), "| gen", dict(gen), "| setting", setting, "| design", design)

# ── 머리말·초록 ─────────────────────────────────────────────
rep("**Draft**: 2026-08-06 (한글본) · corpus 98 studies", f"**Draft**: 2026-09-18 (한글본) · corpus {N} studies")
rep("인용 추적과 보조 색인(supplementary index)을 더해 98편의 적격 연구를 확인했다.",
    f"인용 추적과 보조 색인(supplementary index)을 더해 {N}편의 적격 연구를 확인했다.")
rep("방향성 레코드 187건 중 74건(40%)이 역방향이었으며,", "방향성 레코드 208건 중 88건(42%)이 역방향이었으며,")

# ── 서론: 자기 결과 수치 삭제 ─────────────────────────────────
rep("이 주제의 문헌은 최근 빠르게 늘었다. 본 리뷰에 포함된 98편 중 70편(71%)이 2020년 이후에 게재됐다. 이러한 증가는",
    "이 주제의 문헌은 최근 빠르게 늘었고, 이러한 증가는")
rep("웨어러블 센서를 이용해 훨씬 큰 규모로 행태를 측정할 수 있게 됐으며, 본 리뷰의 코퍼스에도 자전거 통행 81,403건이나 가로 구간 13,322개를 분석한 연구가 포함돼 있다.",
    "웨어러블 센서를 이용해 훨씬 큰 규모로 행태를 측정할 수 있게 됐다.")
rep("다만 이러한 변화에는 방법론적 한계도 따른다. 새로 늘어난 연구의 상당수는 관찰 연구이거나 대규모 자료를 이용한 연구로, 변수 간 연관은 보여 줄 수 있지만 인과관계를 확인하기는 어렵다. 실제로 본 리뷰의 코퍼스에서 현장실험이나 자연실험(field or natural experiment) 설계를 사용한 연구는 19%에 그쳤다(3.2절). 자료의 양은 크게 늘었지만, 음환경을 직접 조작(manipulation)해 행태 변화를 확인한 연구는 여전히 부족하다.",
    "다만 새로 늘어난 연구의 상당수는 관찰 연구이거나 대규모 자료를 이용한 연구여서, 변수 간 연관은 보여 줄 수 있지만 음환경을 직접 조작(manipulation)하지 않는 한 인과관계를 확인하기는 어렵다.")
rep("그러나 선행 문헌을 검토하는 과정에서 반대 방향, 즉 사람의 활동과 점유(occupancy)가 음환경을 형성하는 경로를 다룬 연구도 적지 않다는 점이 확인됐고, 이에 따라",
    "그러나 프로토콜 등록 전 예비 검색에서 반대 방향, 즉 사람의 활동과 점유(occupancy)가 음환경을 형성하는 경로를 다룬 연구도 적지 않다는 점을 확인했고, 이에 따라")
rep("본 리뷰의 코퍼스에서 방향을 판별할 수 있었던 레코드 187건 가운데 74건(40%)이 이러한 역방향이었다(3.6절). 역방향 레코드는 행태 도메인에 따라 고르게 분포하지 않았다. 실험적으로 조작하기 쉬운 이동과 체류에서는 순방향 설계가 대부분이었지만, 공간이용, 활동, 사회적 행태에서는 순방향과 역방향 레코드가 비슷한 비중을 보였다. 이는 집합적이고 지속적인 활동이 음환경을 만드는 요인으로 이미 연구되고 있지만, 분야의 일반적인 틀에서는 상대적으로 덜 주목받아 왔음을 시사한다.",
    "이러한 연구는 음향학, 관광, 공중보건, 도시 분석 분야에 흩어져 있어 순방향 연구와 함께 종합된 적이 없다.")

# ── 3.1 ───────────────────────────────────────────────────
a = t.index("데이터베이스 검색에서 2,073건(Web of Science 1,010건")
b = t.index("선별 과정은 Fig. 1에 제시했다.")
t = t[:a] + (
    "데이터베이스 검색에서 2,073건(Web of Science 1,010건, Scopus 850건, PubMed 213건)을 확인했다. 중복 757건을 제거한 1,316건을 스크리닝해 1,127건을 배제했고, "
    "189건을 전문 확보 대상으로 선정했다. 이 가운데 6건은 전문을 확보하지 못했다. 전문을 평가한 183편 가운데 86편을 배제하고 6편을 민감도 분석용으로 분류해, "
    "데이터베이스 경로에서 91편을 포함했다. 첫 선별 시점(2026년 8월)에 데이터베이스 경로에서 포함되거나 민감도 분석용으로 분류된 84편의 인용을 추적해 "
    "기존 풀에 없던 2,073건을 추가로 확인했다. 제목 기준 자동 선별 후 428건을 제목으로, 146건을 초록으로 스크리닝해 79건을 전문 확보 대상으로 선정했고, "
    "확보한 74편 가운데 54편을 배제하고 2편을 민감도 분석용으로 분류해 18편을 추가로 포함했다. 보조 색인 경로에서는 329건을 스크리닝해 15건을 전문 확보 대상으로 "
    "선정했고, 학술지 논문이 아닌 2건을 제외한 13편을 평가해 4편을 포함하고 1편을 민감도 분석용으로 분류했다. "
    f"최종적으로 {N}편을 리뷰에 포함했으며, 민감도 분석에만 사용한 9편을 더하면 분석 대상은 122편이다. ") + t[b:]

# ── 3.2 ───────────────────────────────────────────────────
rep("포함된 98편 중 70편(71%)이 2020년 이후에 게재됐으며, 출판연도의 중앙값은 2022년이고 가장 오래된 연구는 1975년에 발표됐다. 참가자 수를 보고한 56편의 중앙값은 124명(IQR 30–400)이었다. 나머지 연구는 가로 구간, GPS 궤적, 스트리트뷰 표본점처럼 참가자가 아닌 단위로 표본을 보고했으며, 가장 큰 표본은 자전거 통행 81,403건이었다. 세팅은 가로(41편)와 공원(32편)이 대부분이었고, 광장(10편), 캠퍼스, 주거 옥외공간, 수변이 뒤를 이었다. 연구 설계는 관찰과 설문을 결합한 혼합 연구(34편)가 가장 많았고, 설문 연구(25편), 현장실험 또는 자연실험(19편, 19%), 현장 관찰(7편), 센서·빅데이터 연구(6편), 실험실·가상현실 실험(5편)이 뒤를 이었다.",
    f"포함된 {N}편 중 {per['2020–2026']}편({pct(per['2020–2026'])})이 2020년 이후에 게재됐으며, 출판연도의 중앙값은 2022년이고 가장 오래된 연구는 1975년에 발표됐다. "
    "참가자 수를 보고한 64편의 중앙값은 130명(IQR 30–408)이었다. 9편은 가로 구간, GPS 궤적, 스트리트뷰 표본점처럼 참가자가 아닌 단위로 표본을 보고했고 "
    "가장 큰 표본은 자전거 통행 81,403건이었으며, 나머지 40편은 표본 크기를 보고하지 않았거나 참가자 수로 환산할 수 없었다. "
    f"세팅은 가로({dict(setting)['Street']}편)와 공원({dict(setting)['Park']}편)이 대부분이었고, 광장({dict(setting)['Square / plaza']}편), 수변({dict(setting)['Waterfront']}편), "
    f"캠퍼스({dict(setting)['Campus']}편), 주거 옥외공간({dict(setting)['Residential open space']}편)이 뒤를 이었다. "
    f"연구 설계는 관찰과 설문을 결합한 혼합 연구({design[0][1]}편)가 가장 많았고, 설문 연구({design[1][1]}편), 현장실험 또는 자연실험({design[2][1]}편, {pct(design[2][1])}), "
    f"현장 관찰({design[3][1]}편), 센서·빅데이터 연구({design[4][1]}편), 실험실·가상현실 실험({design[5][1]}편)이 뒤를 이었다.")
rep("연구가 수행된 국가는 중국이 38편(39%)으로 가장 많았고, 스페인(7편), 영국(6편), 독일(5편)이 뒤를 이었으며, 4편은 자료 수집 국가를 보고하지 않았다(Fig. 2a). 역방향 연구 33편 가운데 31편은 2016년 이후에 발표됐다(Fig. 2b).",
    f"연구가 수행된 국가는 중국이 {cc['China']}편({pct(cc['China'])})으로 가장 많았고, 영국({cc['United Kingdom']}편), 독일({cc['Germany']}편), 스페인({cc['Spain']}편), "
    f"이탈리아({cc['Italy']}편)가 뒤를 이었으며, {cc['Not reported']}편은 자료 수집 국가를 보고하지 않았다(Fig. 2a). "
    f"역방향 연구 {len(rev)}편 가운데 {sum(int(r['year']) >= 2016 for r in rev)}편은 2016년 이후에 발표됐다(Fig. 2b).")

# ── Table 1 (사용자 서식 기준본의 행 구성 유지, 수치만 정본으로) ─────────────
tb0 = t.index("**Table 1.** Characteristics of the included studies (n = 98).")
tb1 = t.index("*Note.* Percentages are of the 98 included studies.")
old = t[tb0:tb1]
assert old.count("| Characteristic | Studies | % |") == 1
listed = ["China", "United Kingdom", "Germany", "Spain", "Italy", "Australia", "Czechia", "Switzerland", "United States"]
others = [k for k in cc if k not in listed and k != "Not reported"]
rows = [f"**Table 1.** Characteristics of the included studies (n = {N}).", "",
        "| Characteristic | Studies | % |", "|---|---|---|", "| **Publication period** | | |"]
rows += [f"| {k} | {per[k]} | {pct(per[k])} |" for k in ("≤2009", "2010–2019", "2020–2026")]
rows += ["| **Study design** | | |"] + [f"| {k} | {v} | {pct(v)} |" for k, v in design]
rows += ["| **Setting** † | | |"] + [f"| {k} | {v} | {pct(v)} |" for k, v in setting]
rows += ["| **Country** † | | |"] + [f"| {k} | {cc[k]} | {pct(cc[k])} |" for k in listed]
rows += [f"| Not reported | {cc['Not reported']} | {pct(cc['Not reported'])} |",
         f"| Other countries (n = {len(others)}) | {sum(cc[k] for k in others)} | — |"]
rows += ["| **Direction of relationship** | | |"] + [f"| {k.capitalize()} | {dirc[k]} | {pct(dirc[k])} |" for k in ("forward", "reverse", "both")]
rows += ["| **Behavioural measurement generation** † | | |"] + [f"| {k} | {gen[k]} | {pct(gen[k])} |" for k in ("G1", "G2", "G3")]
rows += ["| **Methodological quality (MMAT 2018)** | | |"] + [f"| {k.capitalize()} | {tier[k]} | {pct(tier[k])} |" for k in ("high", "moderate", "low")]
# 기존 블록이 표 뒤에 두던 빈 줄 수를 그대로 유지
tail = old[len(old.rstrip("\n")):]
t = t[:tb0] + "\n".join(rows) + tail + t[tb1:]
rep("*Note.* Percentages are of the 98 included studies.", f"*Note.* Percentages are of the {N} included studies.")

# ── 3.3 ───────────────────────────────────────────────────
rep("포함된 98편의 MMAT 등급은 high 22편(22%), moderate 40편(41%), low 36편(37%)이었다(Fig. 3). MMAT 범주는 정량 기술 연구 51편, 비무작위 정량 연구 22편, 혼합 연구 15편, 질적 연구 7편, 무작위배정 연구 2편이었고, 1편은 범주를 판정할 수 없었다.",
    f"포함된 {N}편의 MMAT 등급은 high {tier['high']}편({pct(tier['high'])}), moderate {tier['moderate']}편({pct(tier['moderate'])}), low {tier['low']}편({pct(tier['low'])})이었다(Fig. 3). "
    "MMAT 범주는 정량 기술 연구 58편, 비무작위 정량 연구 25편, 혼합 연구 16편, 질적 연구 11편, 무작위배정 연구 2편이었고, 1편은 범주를 판정할 수 없었다.")
rep("정량 기술 연구 51편 중 표본 대표성을 충족한 연구는 10편(20%)이었고 24편은 미충족, 17편은 판단불가로 판정됐다. 무응답 편의가 낮다고 판정된 연구는 16편(31%)이었으며, 14편은 미충족, 21편은 판단불가였다. 비무작위 정량 연구 22편 중 참가자가 목표 모집단을 대표한다고 판정된 연구는 4편(18%)이었고 9편은 미충족, 9편은 판단불가였으며, 교란을 적절히 통제한 연구는 9편(41%)이었고 7편은 미충족, 6편은 판단불가였다.",
    "정량 기술 연구 58편 중 표본 대표성을 충족한 연구는 10편(17%)이었고 28편은 미충족, 20편은 판단불가로 판정됐다. 무응답 편의가 낮다고 판정된 연구는 17편(29%)이었으며, "
    "14편은 미충족, 27편은 판단불가였다. 비무작위 정량 연구 25편 중 참가자가 목표 모집단을 대표한다고 판정된 연구는 6편(24%)이었고 10편은 미충족, 9편은 판단불가였으며, "
    "교란을 적절히 통제한 연구는 10편(40%)이었고 9편은 미충족, 6편은 판단불가였다.")
rep("어떤 기준도 충족하지 못한 연구는 4편이었다.", "어떤 기준도 충족하지 못한 연구는 5편이었다.")

# ── 3.6·3.7 ───────────────────────────────────────────────
rep("98편에서 추출한 행태 도메인 수준 레코드 212건 가운데 113건은 순방향(소리 → 행태), 74건은 역방향(행태 → 소리), 25건은 양방향이었다. 순방향 또는 역방향으로 분류된 187건 중 역방향은 74건(40%)이었고, 양방향 레코드까지 분모에 포함하면 35%였다.",
    f"{N}편에서 추출한 행태 도메인 수준 레코드 248건 가운데 120건은 순방향(소리 → 행태), 88건은 역방향(행태 → 소리), 40건은 양방향이었다. "
    "순방향 또는 역방향으로 분류된 208건 중 역방향은 88건(42%)이었고, 양방향 레코드까지 분모에 포함하면 35%였다.")
rep("이동(순방향 33건, 역방향 8건)과 체류(16건, 8건)에서는 순방향 레코드가 대부분이었지만, 공간이용(26건, 23건), 활동(18건, 21건), 사회적 행태(20건, 14건)에서는 두 방향의 레코드 수가 비슷했다.",
    "이동(순방향 35건, 역방향 9건)과 체류(18건, 12건)에서는 순방향 레코드가 많았지만, 공간이용(27건, 28건), 활동(19건, 24건), 사회적 행태(21건, 15건)에서는 두 방향의 레코드 수가 비슷했다.")
rep("센서, GPS, 영상, 빅데이터를 이용한 행태 측정(G3)은 2010–2019년 4편에서 2020년 이후 20편으로 늘었고, 같은 기간 자기보고(G1)는 14편에서 39편으로, 체계적 관찰(G2)은 11편에서 28편으로 늘었다(Fig. 6). 22편은 두 세대 이상의 방법을 함께 사용했으며, 이러한 연구는 2010년대 5편에서 2020년 이후 17편으로 늘었다.",
    "센서, GPS, 영상, 빅데이터를 이용한 행태 측정(G3)은 2010–2019년 4편에서 2020년 이후 27편으로 늘었고, 같은 기간 자기보고(G1)는 17편에서 46편으로, "
    "체계적 관찰(G2)은 12편에서 33편으로 늘었다(Fig. 6). 30편은 두 세대 이상의 방법을 함께 사용했으며, 이러한 연구는 2010년대 6편에서 2020년 이후 23편으로 늘었다.")

# ── 4.6 한계 ───────────────────────────────────────────────
rep("포함 연구의 39%가 중국에서 수행됐고 71%가 2020년 이후에 게재됐다.",
    f"포함 연구의 {pct(cc['China'])}가 중국에서 수행됐고 {pct(per['2020–2026'])}가 2020년 이후에 게재됐다.")
a = t.index("또한 세 경로를 통틀어 122건의 전문을 확보하지 못했다")
b = t.index("\n\n", a)
t = t[:a] + ("또한 전문 확보 대상 283건 가운데 11건은 구독 전용이거나 원문을 찾지 못해 평가하지 못했다(데이터베이스 경로 6건, 인용 추적 5건). "
             "인용 추적 레코드에 적용한 제목 기준 자동 선별은 제목에 행태 어휘가 없는 적격 연구를 놓쳤을 수 있으며, 제목만으로 선별하면 제목과 초록을 함께 볼 때보다 "
             "민감도가 낮다는 보고가 있다[55].") + t[b:]

# ── 결론 ─────────────────────────────────────────────────
rep("방향성 레코드 187건 중 74건(40%)이 행태나 활동에서 음환경 또는 사운드스케이프로 향했다.",
    "방향성 레코드 208건 중 88건(42%)이 행태나 활동에서 음환경 또는 사운드스케이프로 향했다.")

open(P, "w", encoding="utf-8", newline="").write(t)
print("OK — 원고 갱신(메타분석 제외)")
