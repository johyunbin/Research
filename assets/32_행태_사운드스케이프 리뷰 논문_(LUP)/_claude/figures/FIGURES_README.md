# Paper32 — Figure 세트 (투고용, 2026-08-03)

생성 = `../make_figures.py` (matplotlib, 300 dpi, PNG+PDF 동시 출력). 라벨 전부 영문(투고 대비).
수치 출처는 각 항목의 근거 파일과 1:1 대응 — 그림만 보고도 재현 가능.

| 파일 | 논문 내 위치 | 내용 | 근거 |
|---|---|---|---|
| **Fig1_PRISMA** | Methods | PRISMA 2020 흐름도 (2,073 → 1,316 → 189 → 100 → 81) | `../fulltext/prisma_flow.md` |
| **Fig2_Forest** | Results | 4개 클러스터 forest plot (보행속도·체류·사회·상관) | `../fulltext/ma/ma_summary.md` |
| **Fig3_EvidenceMap** | Results | 행태 도메인 × 음원 히트맵 (5×6) | `../fulltext/evidence_map.md` |
| **Fig4_Direction** | Results | 방향(forward/reverse/both) × 도메인 누적막대 | 동일 |
| **Fig5_Methods** | Results/Discussion | 측정방법 세대 × 시기 (RQ3) | 동일 |
| **Fig6_Framework** | Discussion(종합) | 개념 프레임워크 — 양방향 루프 + 관여경사 + 측정3세대 | `../fulltext/ma/ma_summary.md`·`quality_summary.md`·`evidence_map.md` |
| **Fig7_GeoTime** | Limitations | 국가 분포 + 연도×방향 | `../fulltext/geo_time_counts.csv` |
| **Fig8_Quality** | Results (RoB) | MMAT 2018 범주별 등급 + 문항별 판정 | `../fulltext/quality_all.csv`·`quality_detail_all.csv` |

Fig1은 `../make_fig1_prisma_v2.py`(2갈래판이 정본 — `make_figures.py`의 구 `fig_prisma()`는 비활성화),
Fig6·Fig7·Fig8은 각각 `../make_fig6_framework.py`, `../make_fig7_geo_time.py`, `../make_fig8_quality.py`.
공통 팔레트 = `../viz_theme.py` (OKLab ΔE·색각이상 시뮬레이션 검증 통과분).

## 그림별 메시지 (캡션 초안)

**Fig 1.** PRISMA 2020 flow diagram with both identification routes. The dominant exclusion reason in
the database branch is animal/wildlife research (n = 479), reflecting the terminological overlap between
soundscape ecology and human soundscape research. The citation-searching branch is shown dashed because
it stops at "reports sought for retrieval" — 71 of 79 reports could not be obtained automatically.

**Fig 2.** Random-effects meta-analytic estimates (REML with Hartung–Knapp adjustment) for four
behavioural clusters. Squares are individual effects (size ∝ inverse variance), diamonds are pooled
estimates. Panels (a)–(c) show consistent directions but confidence intervals spanning zero; only the
correlational cluster (d) reaches significance (r = 0.43, p = 0.005).

**Fig 3.** Evidence map. Every domain × source cell is populated except aircraft noise (n = 2), which
is the clearest gap given the size of the aircraft-noise health literature.

**Fig 4.** Direction of the studied relationship. Movement and staying are dominated by forward
designs, whereas space use, activity and social behaviour show near-parity between forward and
reverse pathways — the empirical basis for a bidirectional framing.

**Fig 5.** Behavioural measurement methods over time. Sensor/GPS/video and big-data measurement grew
from 4 studies (2010–2019) to 17 (2020–), while self-report continued to grow — methods accumulate in
layers rather than replacing one another.

**Fig 6.** Conceptual framework. The loop is bidirectional by construction: the forward path runs
acoustic environment → appraisal → behaviour, the reverse path runs activity → sound production →
acoustic environment. Behaviour is unfolded as an engagement gradient, and each band carries the
quantitative evidence attached to it. The framework's own diagnostic value is visible in the leftmost
band — the avoidance/walking-speed evidence, the most frequently cited claim in this literature, rests
entirely on studies that MMAT rates low.

**Fig 7.** Geographic and temporal distribution. Chinese studies account for 34 of 84 (40%), and 71%
of the corpus appeared after 2020 — the evidence base is both young and geographically narrow, which
bounds how far the pooled estimates travel. Panel (b) also shows the reverse-direction studies
emerging only after 2016.

**Fig 8.** MMAT 2018 appraisal. Two findings drive the Discussion. First, the two randomised studies
score lowest, not highest — their reports omit randomisation, baseline comparability and blinding
entirely, so the design cannot be credited. Second, the largest deficits are *reporting* deficits:
sample representativeness (4.2, only 20% clearly met), nonresponse bias (4.4, 41%) and confounding
(3.4, 37%) fail mostly because the information is absent, not because the studies are known to be
biased. A field that measures sound to three decimal places routinely omits its response rate.

## 재생성

```bash
python make_figures.py && python make_fig1_prisma_v2.py && python make_fig6_framework.py && python make_fig7_geo_time.py && python make_fig8_quality.py
```
수치를 바꾸려면 `make_figures.py` 상단의 각 함수 내 하드코딩 값을 근거 파일과 함께 갱신할 것
(현재 값은 `ma_summary.md`·`evidence_map.md`·`prisma_flow.md`와 일치 확인됨).
