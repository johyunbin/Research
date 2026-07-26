# -*- coding: utf-8 -*-
# 수도권/서울 사회적 거리두기 준연속 재코딩 -> 일별 테이블.
# ordinal '단계' 대신 (영업종료시각 close_hour, 허용 모임인원 gather_limit)로 재코딩(프로토콜 §4).
# 거리두기 = 보조/scaffold dose (primary = 생활인구). 주요 국면전환 앵커는 잘 문서화됨.
# !! 정밀 일자경계·세부수치는 질병청 HWPX(공공데이터포털 15106451) 원본 대조 필요 (open item).
import csv, os
from datetime import date, timedelta

ROOT = r"T:\00_BACKUP\01_한양대학교\00_연구실\00_프로젝트\★_논문화 프로젝트\31_환경소음_이동성"
OUT = os.path.join(ROOT, r"data\processed\distancing_daily_2020-2023.csv")

# (시작일, 영업종료시각[24=무제한], 허용모임인원[999=무제한], 라벨, approx[1=pre-tier 근사])
# 수도권 기준. close_hour=식당/카페 등 다중이용시설 영업제한 시각.
INTERVALS = [
    ("2020-01-01", 24, 999, "정상기(pre-COVID)", 0),
    ("2020-02-29", 24, 999, "주의(대구확산·수도권권고)", 1),
    ("2020-03-22", 21,  10, "강력한 사회적 거리두기(1차)", 1),
    ("2020-04-20", 22,  30, "강력거리두기 완화", 1),
    ("2020-05-06", 24, 999, "생활 속 거리두기(생활방역)", 0),
    ("2020-06-28", 24, 999, "3단계체계 1단계(수도권)", 0),
    ("2020-08-16", 21, 100, "수도권 2단계", 0),
    ("2020-08-30", 21,  50, "수도권 강화2단계(영업21시 첫 도입)", 0),
    ("2020-09-14", 22, 100, "수도권 2단계 완화", 0),
    ("2020-10-12", 24, 999, "수도권 1단계", 0),
    ("2020-11-19", 22, 100, "수도권 1.5단계", 0),
    ("2020-11-24", 21,  50, "수도권 2단계", 0),
    ("2020-12-08", 21,  50, "수도권 2.5단계", 0),
    ("2020-12-23", 21,   4, "수도권 2.5단계+5인이상모임금지", 0),
    ("2021-02-15", 22,   4, "수도권 2단계+5인금지", 0),
    ("2021-07-01", 24,   8, "4단계체계 시행초기(수도권)", 0),
    ("2021-07-12", 22,   4, "수도권 4단계(18시후 2인)", 0),
    ("2021-11-01", 24,  10, "단계적 일상회복(위드코로나)", 0),
    ("2021-12-06", 24,   6, "방역강화(6명)", 0),
    ("2021-12-18", 21,   4, "특별방역대책(4명·21시)", 0),
    ("2022-02-19", 22,   6, "사적모임6명·22시", 0),
    ("2022-03-05", 23,   6, "23시 완화", 0),
    ("2022-03-21", 23,   8, "8명·23시", 0),
    ("2022-04-04", 24,  10, "10명·24시(해제직전)", 0),
    ("2022-04-18", 24, 999, "거리두기 전면해제", 0),
]
# 2022-04-18 이후 ~ 2023-12-31 = 거리두기 없음(정상기). 2023 전체 무제한.

def d(s):
    y, m, dd = map(int, s.split("-"))
    return date(y, m, dd)

iv = [(d(s), ch, gl, lab, ap) for s, ch, gl, lab, ap in INTERVALS]

# 연속 stringency index: 영업제한(24-close_hour, 0~3) + 모임제한( log-ish 점수 )
def gather_score(gl):
    if gl >= 999: return 0
    return {10: 1, 8: 1.5, 6: 2, 5: 2.5, 4: 3, 2: 4}.get(gl, 2)

rows = []
cur = date(2020, 1, 1); end = date(2023, 12, 31)
while cur <= end:
    # 가장 최근 시작일 구간
    seg = iv[0]
    for s in iv:
        if s[0] <= cur: seg = s
        else: break
    _, ch, gl, lab, ap = seg
    close_strict = 24 - ch            # 0(무제한)~3(21시)
    gs = gather_score(gl)
    stringency = round(close_strict + gs, 2)   # 0(정상)~ ~7(최강)
    any_dist = 1 if (ch < 24 or gl < 999) else 0
    rows.append([cur.isoformat(), lab, ch, (gl if gl < 999 else ""), close_strict,
                 gs, stringency, any_dist, ap])
    cur += timedelta(days=1)

with open(OUT, "w", newline="", encoding="utf-8-sig") as f:
    w = csv.writer(f)
    w.writerow(["date", "phase_label", "close_hour", "gather_limit",
                "close_strict", "gather_score", "stringency", "any_distancing", "approx"])
    w.writerows(rows)

print(f"일수: {len(rows)} ({rows[0][0]} ~ {rows[-1][0]})")
print(f"-> {OUT}")
# 앵커 점검
import collections
phase_days = collections.Counter(r[1] for r in rows)
print("\n주요 국면 일수:")
for lab, n in sorted(phase_days.items(), key=lambda x: -x[1])[:8]:
    print(f"  {n:>4}d  {lab}")
print("\nstringency 전환 샘플(주요 앵커):")
for probe in ["2020-01-15","2020-03-25","2020-09-01","2020-12-25","2021-07-15","2021-11-05","2021-12-20","2022-04-20","2023-06-01"]:
    r = next(x for x in rows if x[0] == probe)
    print(f"  {probe}: stringency={r[6]} close={r[2]}시 gather={r[3] or '무제한'} | {r[1]}")
