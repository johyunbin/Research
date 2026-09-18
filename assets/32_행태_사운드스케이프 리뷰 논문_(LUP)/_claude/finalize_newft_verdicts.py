# -*- coding: utf-8 -*-
"""
Paper32 — 추가 전문 77건 판정 확정 (1차 평가 + 블라인드 재판정 36건 + 불일치 조정) · 2026-09-18

입력: fulltext/newft_combined/{verdict,extract,mmat,detail,es}.csv (1차) · fulltext/newft_recheck/rc_*_verdict.csv (재판정)
규칙:
  - 재판정하지 않은 41건 = 1차 판정 유지
  - 1차·재판정 결론(포함/SENS/배제) 일치 = 1차 판정 유지. 둘 다 배제인데 코드가 다르면 번호가 앞선 코드(등록본: 처음 해당하는 기준)
  - 결론 불일치 3건 = 아래 ADJUDICATION(원문 근거)
  - 최종 배제가 된 레코드는 extract·mmat·detail·es 에서 제거
출력: fulltext/newft_final/{verdict,extract,mmat,detail,es}.csv · fulltext/newft_final/adjudication_log.md · 재판정 일치도 출력
"""
import csv, glob, os, sys
from collections import Counter

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
TAG = next((a.split("=", 1)[1] for a in sys.argv if a.startswith("--tag=")), "")   # "" = 1차(9/17 확보분), "b" = 2차(9/18 확보분)
COMB = os.path.join(FT, "newft_combined" + (f"_{TAG}" if TAG else ""))
RECHECK_GLOB = os.path.join(FT, "newft_recheck" + (f"_{TAG}" if TAG else ""), ("rcb" if TAG else "rc") + "_*_verdict.csv")
OUT = os.path.join(FT, "newft_final" + (f"_{TAG}" if TAG else ""))

ADJUDICATION_B = {
    "137": ("FINAL_INCLUDE", "", "R3", "low",
            "조정(1차 포함 R3 low ↔ 재판정 SENS_ONLY R2 low): 이 연구의 노출은 야간 하이킹이라는 활동이고 결과 주제 하나가 "
            "'Unique Soundscapes and Night Sky'(사운드스케이프 지각·평가)다. 활동 유형이 청각 지각을 좌우하는 구조라 R3(역방향) "
            "선례 726(단독 대 가이드 산림보행 → 청각요소 지각)·761·804 와 같다. 재참여 의향은 부수 결과이므로 R2 를 적용하지 않는다. "
            "다만 활동 간 비교가 없고 행태 측정이 없어 근거는 약하다(confidence low)."),
    "1074": ("FINAL_EXCLUDE", "X4", "", "medium",
             "조정(1차 포함 R3 low ↔ 재판정 배제 X4): 수탉 울음 가청 지점 지도는 음환경 산출물이지만, 방문객 결과는 인상·태도이고 "
             "급이·유기·창문 닫기 같은 행태는 정보원 인용·서술일 뿐 코딩·분석되지 않았다. 행태→음환경 경로도 저자 서술이다. "
             "행태가 결과로 분석된 1190(경로 선택)·1191(체류·촬영 기록)과 달라 배제한다."),
}
ADJUDICATION = {
    "186": ("FINAL_EXCLUDE", "X3", "", "high",
            "조정(1차 포함 low ↔ 재판정 배제 X3): 운전 행태에 따른 차량 엔진룸·후륜 근접장 방출음만 측정했고 가로·공공공간 음환경은 "
            "측정하지 않았다. 같은 주제의 194(Ibarra 2012)가 1차·재판정 모두 X3 이므로 일관되게 배제. 경적 선례 799·1138 은 도로변 "
            "소음계로 가로 음환경을 측정해 구조가 다르다."),
    "624": ("FINAL_EXCLUDE", "X4", "", "medium",
            "조정(1차 포함 R3 low ↔ 재판정 배제 X4): 여행기 근거이론 분석에서 음풍경(Table 4)과 활동(Table 2)이 따로 코딩됐을 뿐 "
            "둘의 관계는 분석되지 않았고, '조용한 곳에서 느린 여가활동'(§5.3·§5.7)은 저자 종합 서술이다. 고요(tranquillity)는 조용함·"
            "사회 안정·경관·지역문화의 총체 경험(§5.2.2)이라 사운드스케이프 평가로 보기 어렵다."),
    "1191": ("FINAL_INCLUDE", "", "", "low",
             "조정(1차 포함 low ↔ 재판정 배제 X4): §2.1 에 카메라·위치기기로 관광객 체류시간·사진 촬영 빈도를 표본 지점별로 기록했다고 "
             "명시했고, §4 에서 자연음 구역의 높은 쾌적·고요 점수와 긴 체류·잦은 촬영의 상관을 보고한다. 행태를 측정했으나 결과 수치를 "
             "보고하지 않은 경우이므로 적격성으로 배제하지 않고(결과 보고 여부로 적격성을 정하지 않음) 보고 결함을 품질평가·서술에 반영한다."),
}
ORDER = ["X1", "X2", "X3", "X4", "X5", "X6", "X7"]


def read(p):
    return list(csv.DictReader(open(p, encoding="utf-8-sig")))


def write(p, rows, keys=None):
    keys = keys or list(dict.fromkeys(k for r in rows for k in r))
    with open(p, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=keys)
        w.writeheader()
        w.writerows(rows)


def binary(v):
    return "exclude" if v == "FINAL_EXCLUDE" else "include"


def main():
    global ADJUDICATION
    if TAG == "b":
        ADJUDICATION = ADJUDICATION_B
    first = {r["id"]: r for r in read(os.path.join(COMB, "verdict.csv"))}
    rc = {}
    for f in sorted(glob.glob(RECHECK_GLOB)):
        for r in read(f):
            rc[r["id"]] = r
    assert set(rc) <= set(first) and set(ADJUDICATION) <= set(rc)
    print(f"1차 {len(first)}건 · 재판정 {len(rc)}건")

    final, log = [], []
    for i, a in first.items():
        row = {k: a[k] for k in ("id", "branch", "verdict", "reason_code", "boundary_rule", "confidence", "rationale")}
        row.update(first_verdict=a["verdict"], recheck_verdict="", decision="1차 유지(재판정 대상 아님)")
        if i in rc:
            r = rc[i]
            row["recheck_verdict"] = r["verdict"]
            if i in ADJUDICATION:
                v, code, rule, conf, why = ADJUDICATION[i]
                row.update(verdict=v, reason_code=code, boundary_rule=rule, confidence=conf,
                           rationale=why, decision="불일치 조정")
                log.append((i, a, r, row))
            elif a["verdict"] == r["verdict"]:
                row["decision"] = "1차·재판정 일치"
                if a["verdict"] == "FINAL_EXCLUDE" and a["reason_code"] != r["reason_code"]:
                    code = min(a["reason_code"], r["reason_code"], key=ORDER.index)
                    if code != a["reason_code"]:
                        row["reason_code"] = code
                        row["rationale"] = r["rationale"]
                        row["decision"] = f"1차·재판정 일치(코드 {a['reason_code']}→{code}: 번호 순서)"
            else:
                raise SystemExit(f"조정 규칙이 없는 불일치: {i} {a['verdict']} vs {r['verdict']}")
        final.append(row)

    keep = {r["id"] for r in final if r["verdict"] != "FINAL_EXCLUDE"}
    os.makedirs(OUT, exist_ok=True)
    write(os.path.join(OUT, "verdict.csv"), final)
    for k in ("extract", "mmat", "detail", "es"):
        rows = [r for r in read(os.path.join(COMB, f"{k}.csv")) if r["id"] in keep]
        write(os.path.join(OUT, f"{k}.csv"), rows)
    ex_ids = {r["id"] for r in read(os.path.join(OUT, "extract.csv"))}
    assert ex_ids == keep, f"추출 누락: {keep - ex_ids}"

    # 재판정 일치도(포함·SENS vs 배제)
    pairs = [(binary(first[i]["verdict"]), binary(rc[i]["verdict"])) for i in rc]
    n = len(pairs)
    po = sum(a == b for a, b in pairs) / n
    ca, cb = Counter(a for a, _ in pairs), Counter(b for _, b in pairs)
    pe = sum(ca[l] * cb[l] for l in ("include", "exclude")) / n ** 2
    kappa = (po - pe) / (1 - pe)

    L = [f"# 추가 전문 {len(first)}건 판정 확정 기록 (2026-09-18{' · 2차 확보분' if TAG else ''})", "",
         f"- 1차 평가 {len(first)}건(원문 전체 읽기) · 블라인드 재판정 {len(rc)}건(포함·SENS + 경계 배제, 1차 결과 비공개)",
         f"- 재판정 결론 일치 {sum(a == b for a, b in pairs)}/{n} ({po:.1%}), Cohen's κ = {kappa:.2f} (포함·SENS vs 배제). "
         "재판정 대상이 경계 사례 위주라 전체 일치도보다 보수적인 값이다.", "",
         "## 결론 불일치 조정 3건", ""]
    for i, a, r, row in log:
        L += [f"### {i}", f"- 1차: {a['verdict']} {a['reason_code']} ({a['confidence']}) — {a['rationale']}",
              f"- 재판정: {r['verdict']} {r['reason_code']} ({r['confidence']}) — {r['rationale']}",
              f"- **확정: {row['verdict']} {row['reason_code']}** — {row['rationale']}", ""]
    L += ["## 코드만 다른 일치 배제", ""]
    L += [f"- {r['id']}: {r['decision']}" for r in final if "코드" in r["decision"]]
    L += ["", "## 최종 집계", "",
          f"- {dict(Counter((r['branch'], r['verdict']) for r in final))}",
          f"- 배제 사유: {dict(Counter(r['reason_code'] for r in final if r['reason_code']))}"]
    open(os.path.join(OUT, "adjudication_log.md"), "w", encoding="utf-8").write("\n".join(L) + "\n")

    print(f"재판정 일치 {sum(a == b for a, b in pairs)}/{n} ({po:.1%}) · κ={kappa:.2f}")
    print("최종:", Counter((r["branch"], r["verdict"]) for r in final))
    print("포함·SENS:", sorted(keep, key=lambda x: (x.startswith(("CT", "OAS")), x.zfill(8))))
    print("배제 사유:", Counter(r["reason_code"] for r in final if r["reason_code"]))


if __name__ == "__main__":
    main()
