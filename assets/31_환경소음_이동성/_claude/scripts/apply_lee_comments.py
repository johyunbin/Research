# -*- coding: utf-8 -*-
# 교신저자(Kyung Tae Lee) 수정방향 메모 6건 반영 (Manuscript_수정 방향.docx의 Word 코멘트).
# c0 초록 도입 교체 / c1 §1.3 왜-지금-코로나 / c2 §3.2 → 4.3절 전방 참조 / c3 §4.5 자료의 현재 가치
# c4 결론 = 공간 재분포 프레이밍 / c5 결론 한계 1문장. EN·KR 정본 동시 적용(문단 재충전·서식 상속).
import os, sys
from docx import Document
sys.path.insert(0, os.path.dirname(__file__))
import docxtools_v9 as T

def para_size(p, default=11.0):
    for r in p.runs:
        if r.font.size is not None:
            return r.font.size.pt
    st = p.style
    while st is not None:
        if st.font.size is not None:
            return st.font.size.pt
        st = st.base_style
    return default

EN = r"C:\Users\wh850\Research\assets\31_환경소음_이동성\01_논문작업\Manuscript_EN_20260727_005544.docx"
KR = r"C:\Users\wh850\Research\assets\31_환경소음_이동성\01_논문작업\Manuscript_20260726_231342.docx"

C1_EN = (" The pandemic itself has ended, but the episode retains a methodological value that ordinary "
         "urban conditions rarely supply: mobility moved in large, graded and reversible steps for "
         "reasons unrelated to the sound environment, so the noise response to mobility can be "
         "separated from the everyday factors that move both together.")
C3_EN = ("Although the distancing era has passed, the graded mobility variation it produced remains a "
         "rare empirical benchmark for judging how future changes in urban activity, including "
         "travel-demand management, are likely to move neighbourhood noise. ")

EDITS_EN = [
 ("How strongly the urban acoustic", [
   ("How strongly the urban acoustic environment responds to human mobility is a first-order question "
    "for mobility, land-use and noise policy, yet mobility rarely varies exogenously.",
    "Exogenous changes in urban mobility are rare, limiting opportunities to link observed population "
    "activity with corresponding changes in daytime and night-time noise."),
 ]),
 ("The COVID-19 pandemic supplied", [
   ("evidence of how broadly confinement reshaped the urban environment.",
    "evidence of how broadly confinement reshaped the urban environment." + C1_EN),
 ]),
 ("Separating day and night", [
   ("day-night difference itself remains statistically unresolved.",
    "day-night difference itself remains statistically unresolved; Section 4.3 weighs three candidate "
    "explanations for the night-time null."),
 ]),
 ("For planning and policy", [
   ("Whether these results generalise to Tokyo,",
    C3_EN + "Whether these results generalise to Tokyo,"),
 ]),
 ("Using Korea's graded distancing", [
   ("in that strict sense they are conditional associations.",
    "in that strict sense they are conditional associations. They also read presence as activity and "
    "average over unobserved differences in sensor installation context."),
   ("For policy, at least at the level of relative neighbourhood change, mobility-reducing urban "
    "policies (car-free streets, fifteen-minute cities) are unlikely on their own to lower urban noise "
    "by much, and noise targets will require source- and propagation-stage interventions alongside them.",
    "For policy, the spatial redistribution of mobility alone is unlikely to deliver large noise "
    "reductions: even lockdown-scale reallocation moved neighbourhood noise by well under a decibel, "
    "and noise targets will require source- and propagation-stage interventions alongside mobility "
    "measures."),
 ]),
]

C1_KR = (" 팬데믹 자체는 끝났지만, 이 시기는 정상적인 도시 환경에서는 좀처럼 주어지지 않는 방법론적 가치를 지닌다. "
         "이동량이 소음 환경과 무관한 이유로 크게·단계적·가역적으로 움직였기에, 이동량과 소음을 함께 움직이는 "
         "일상적 요인들로부터 이동량 변화에 대한 소음 반응을 분리할 수 있다.")
C3_KR = ("거리두기 시대는 지났지만, 그 시기가 만든 단계적 이동량 변동은 교통수요관리를 포함한 향후 도시 활동 "
         "변화가 지역 소음을 얼마나 움직일지를 가늠할 드문 실증 기준으로 남는다. ")

EDITS_KR = [
 ("How strongly the urban acoustic", EDITS_EN[0][1]),
 ("코로나-19 대유행은", [
   ("봉쇄가 도시 환경 전반에 미친 충격을 보였다.",
    "봉쇄가 도시 환경 전반에 미친 충격을 보였다." + C1_KR),
 ]),
 ("주간과 야간을 나눠 보면", [
   ("주야 효과의 차이 자체는 통계적으로 확립되지 않는다.",
    "주야 효과의 차이 자체는 통계적으로 확립되지 않는다. 야간 비유의에 대한 세 가지 대안 설명은 4.3절에서 검토한다."),
 ]),
 ("계획·정책 측면에서", [
   ("서울은 세계 최고 밀도의 메가시티 중 하나이며,",
    C3_KR + "서울은 세계 최고 밀도의 메가시티 중 하나이며,"),
 ]),
 ("우리는 한국의 단계적 거리두기를 자연실험", [
   ("보수적 비교에서 추정된 조건부 연관이다.",
    "보수적 비교에서 추정된 조건부 연관이다. 또한 이 추정치는 체류(presence)를 활동으로 읽으며, 관측되지 "
    "않는 센서 설치 맥락의 차이를 평균한 값이다."),
   ("정책적으로 이는, 적어도 동 단위 상대 변화의 수준에서는 이동수요를 줄이는 도시정책(차 없는 거리·15분 "
    "도시)만으로 도시소음을 크게 낮추기 어려울 가능성이 높으며, 소음 목표 달성에는 노면 포장·저소음 차량·차폐 "
    "같은 음원·전파 단계 개입이 병행되어야 함을 시사한다.",
    "정책적으로 이는, 이동량의 공간적 재분포만으로는 큰 폭의 소음 저감을 달성하기 어려움을 시사한다. 봉쇄급 "
    "재배치조차 동 단위 소음을 1 dB에 훨씬 못 미치게 움직였을 뿐이며, 소음 목표 달성에는 노면 포장·저소음 "
    "차량·차폐 같은 음원·전파 단계 개입이 이동량 정책과 병행되어야 한다."),
 ]),
]

for path, edits in ((EN, EDITS_EN), (KR, EDITS_KR)):
    doc = Document(path)
    for marker, subs in edits:
        hits = [p for p in doc.paragraphs if p.text.strip().startswith(marker)]
        assert len(hits) == 1, f"{os.path.basename(path)}: 마커 {len(hits)}건 — {marker}"
        p = hits[0]
        text = p.text
        for old, new in subs:
            assert old in text, f"{os.path.basename(path)}: 원문 미발견 — {old[:60]}"
            text = text.replace(old, new)
        T.fill_para(p, text, size=para_size(p))
    doc.save(path)
    print("OK:", os.path.basename(path))
