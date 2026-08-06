"""design_matrix_v2.csv 정합 검증.

검사 축:
  1. 헤더가 1차(design_matrix.csv)와 동일한가
  2. evidence_studies의 uid가 corpus_v3_extraction.csv에 실재하고 FINAL_INCLUDE인가
  3. quality_mix가 quality_v2.csv 실집계와 일치하는가
  4. n_studies가 uid 개수와 일치하는가 / 레버 내 uid 중복 없음
  5. 동일 표본 쌍(CT0126-CT0322, CT0007-761)이 같은 레버에 이중계상되지 않았는가
  6. confidence 어휘 유효 / direction 어휘 유효
  7. caveat 25단어 이내 / 셀 내 줄바꿈 없음 / UTF-8 디코드 가능
  8. 레버 수 5~10

usage: python verify_design_matrix_v2.py
"""
import csv
import io
import os
import sys
from collections import Counter

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "fulltext")
V1 = os.path.join(BASE, "design_matrix.csv")
V2 = os.path.join(BASE, "design_matrix_v2.csv")
EXT = os.path.join(BASE, "corpus_v3_extraction.csv")
QUA = os.path.join(BASE, "quality_v2.csv")

CONF = {"high", "moderate", "low", "very low"}
DIRE = {"promotes", "inhibits", "mixed"}
SAME_SAMPLE = [("CT0126", "CT0322"), ("CT0007", "761")]

fails, warns = [], []


def load(path):
    with io.open(path, encoding="utf-8-sig") as f:
        return list(csv.DictReader(f))


def main():
    raw = io.open(V2, "rb").read()
    try:
        raw.decode("utf-8")
    except UnicodeDecodeError as e:
        fails.append("UTF-8 디코드 실패: %s" % e)

    v1, v2 = load(V1), load(V2)
    ext = {r["uid"]: r for r in load(EXT)}
    qua = {r["uid"]: r["quality_tier"] for r in load(QUA)}

    if list(v1[0].keys()) != list(v2[0].keys()):
        fails.append("헤더 불일치\n  v1=%s\n  v2=%s" % (list(v1[0]), list(v2[0])))
    else:
        print("[OK] 헤더 1차와 동일 (%d열)" % len(v1[0]))

    if not (5 <= len(v2) <= 10):
        fails.append("레버 수 %d — 5~10 범위 밖" % len(v2))
    print("[OK] 레버 %d개 (1차 %d개)" % (len(v2), len(v1)))

    for r in v2:
        lid = r["lever_id"]
        ids = [x.strip() for x in r["evidence_studies"].split(";") if x.strip()]

        for u in ids:
            if u not in ext:
                fails.append("%s: uid %s 가 corpus_v3_extraction.csv에 없음" % (lid, u))
            elif ext[u]["final_verdict"] != "FINAL_INCLUDE":
                fails.append("%s: uid %s 는 %s (FINAL_INCLUDE 아님)"
                             % (lid, u, ext[u]["final_verdict"]))
            if u not in qua:
                fails.append("%s: uid %s 가 quality_v2.csv에 없음" % (lid, u))

        dup = [u for u, c in Counter(ids).items() if c > 1]
        if dup:
            fails.append("%s: 레버 내 uid 중복 %s" % (lid, dup))

        if str(len(ids)) != r["n_studies"].strip():
            fails.append("%s: n_studies=%s 인데 uid %d개" % (lid, r["n_studies"], len(ids)))

        c = Counter(qua.get(u, "MISSING") for u in ids)
        mix = " · ".join("%s %d" % (k.replace("moderate", "mod"), c[k])
                         for k in ("high", "moderate", "low") if c.get(k))
        if mix != r["quality_mix"].strip():
            fails.append("%s: quality_mix 표기 '%s' vs 실집계 '%s'"
                         % (lid, r["quality_mix"], mix))

        for a, b in SAME_SAMPLE:
            if a in ids and b in ids:
                fails.append("%s: 동일 표본 %s·%s 이중계상" % (lid, a, b))

        if r["confidence"].strip() not in CONF:
            fails.append("%s: confidence 어휘 '%s'" % (lid, r["confidence"]))
        if r["direction"].strip() not in DIRE:
            fails.append("%s: direction 어휘 '%s'" % (lid, r["direction"]))

        nw = len(r["caveat"].split())
        if nw > 25:
            fails.append("%s: caveat %d단어 (25 초과)" % (lid, nw))

        for k, v in r.items():
            if "\n" in (v or "") or "\r" in (v or ""):
                fails.append("%s: 셀 '%s'에 줄바꿈" % (lid, k))

    # 표 출력
    print("\n%-5s %-4s %-24s %-11s %s" % ("lever", "n", "quality_mix(실집계)", "conf", "evidence"))
    for r in v2:
        ids = [x.strip() for x in r["evidence_studies"].split(";") if x.strip()]
        c = Counter(qua.get(u, "?") for u in ids)
        mix = " · ".join("%s %d" % (k.replace("moderate", "mod"), c[k])
                         for k in ("high", "moderate", "low") if c.get(k))
        print("%-5s %-4d %-24s %-11s %s"
              % (r["lever_id"], len(ids), mix, r["confidence"], ";".join(ids)))

    # 코퍼스 커버리지
    cited = set()
    for r in v2:
        cited |= {x.strip() for x in r["evidence_studies"].split(";") if x.strip()}
    inc = [u for u, r in ext.items() if r["final_verdict"] == "FINAL_INCLUDE"]
    print("\n레버 근거로 인용된 연구 %d편 / 포함 %d편 (미인용 %d편)"
          % (len(cited), len(inc), len(inc) - len(cited)))
    ct_cited = sorted(u for u in cited if u.startswith("CT"))
    print("인용추적분 인용 %d/%d: %s" % (len(ct_cited),
          sum(1 for u in inc if u.startswith("CT")), ", ".join(ct_cited)))

    print()
    for w in warns:
        print("[WARN]", w)
    if fails:
        for f in fails:
            print("[FAIL]", f)
        print("\n=== %d건 실패 ===" % len(fails))
        return 1
    print("=== 전 항목 통과 ===")
    return 0


if __name__ == "__main__":
    sys.exit(main())
