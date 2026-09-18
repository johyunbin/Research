# -*- coding: utf-8 -*-
"""
Paper32 — 9월 재판정으로 배제된 uid 1226(Dzhambov et al. 2026, 불가리아 거주지 소음 × 설문 신체활동, R1·X2)의
본문 인용 정리 (2026-09-18)

4.3절의 신체활동 문장은 1226([49])과 951(Huang et al. 2023, [50])을 "가장 규모가 큰 관찰 자료"로 인용했다.
1226 이 코퍼스에서 빠졌으므로 코퍼스 포함 연구 951 만 인용하고, 951 의 실제 자료(헬싱키 가로 13,322 구간의
Strava 달리기 강도, corpus_v4_extraction.csv)를 문장에 적는다. "가장 규모가 큰"은 코퍼스 전체와 대조하지 않은
최상급이라 쓰지 않는다. 같은 문단의 서수 나열("첫째/둘째")도 함께 푼다(메모리 feedback-manuscript-no-ai-tone).
번호 이동과 참고문헌 목록 정리는 insert_citations.py 가 한다(본문에 등장하지 않는 서지는 목록에서 빠진다).
"""
import os, sys

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
P = os.path.join(os.path.dirname(BASE), "01_논문작업", "Manuscript_KO.md")
t = open(P, encoding="utf-8").read()


def rep(a, b):
    global t
    n = t.count(a)
    if n != 1:
        sys.exit(f"STOP count={n}: {a[:80]}")
    t = t.replace(a, b)


rep("흔히 제시되는 두 가지 처방도 본 리뷰의 행태 코퍼스로는 지지되지 않았다. 첫째, 수경시설로 교통소음을 마스킹(masking)해",
    "흔히 제시되는 두 가지 처방도 본 리뷰의 행태 코퍼스로는 지지되지 않았다. 수경시설로 교통소음을 마스킹(masking)해")
rep("둘째, 소음 저감을 신체활동을 늘리는 보편적인 방법으로 보는 관점은 가장 규모가 큰 관찰 자료에서 지지되지 않았다. "
    "활동이 많은 장소는 음향적으로도 붐비는 경우가 많아, 이러한 자료에서는 소음과 활동이 오히려 정적으로 연관될 수 있다[49,50]. "
    "따라서 역인과(reverse causation)와 맥락적 교란(contextual confounding)을 함께 고려해야 한다.",
    "소음 저감을 신체활동을 늘리는 보편적인 방법으로 보는 관점도 코퍼스의 관찰 자료로는 뒷받침되지 않았다. "
    "헬싱키 가로 13,322개 구간의 달리기 기록을 분석한 연구에서는 교통소음이 높은 구간에서 달리기 강도가 오히려 높았다[50]. "
    "활동이 많은 장소는 음향적으로도 붐비는 경우가 많아 이러한 자료에서는 소음과 활동이 정적으로 연관될 수 있으므로, "
    "역인과(reverse causation)와 맥락적 교란(contextual confounding)을 함께 고려해야 한다.")
import re
cited = {int(x) for tok in re.findall(r"\[\d+(?:,\d+)*\]", t.split("## References")[0]) for x in tok[1:-1].split(",")}
assert 49 not in cited, "본문에 [49] 인용이 남아 있다"
open(P, "w", encoding="utf-8", newline="").write(t)
print("OK — 1226 인용 제거(본문). 다음: insert_citations.py 로 재번호")
