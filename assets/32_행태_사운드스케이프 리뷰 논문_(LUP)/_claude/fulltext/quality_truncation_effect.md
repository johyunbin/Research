# Paper32 — 텍스트 절단이 품질평가에 미친 영향 (자체 감사)

인용추적 갈래 1차 품질평가는 논문당 46,000자로 잘린 텍스트로 수행됐다. MMAT는 *'보고가 없으면 CT(can't tell)'* 로 판정하는 도구이므로, **추출 한도가 그대로 품질 점수를 깎는 인공물**이 된다. 포함 16편 중 **13편이 한도를 초과**했다(최장 96,472자).

절단 없는 전문으로 재평가한 뒤 두 결과를 대조했다. **전문 기준 재평가가 정본**이다.

## 판정이 바뀐 논문 — 7편

| 논문 | 등급 | n_yes | 범주 | 바뀐 문항 |
|---|---|---|---|---|
| CT0090 | low → **high** | 2 → 4 | Quantitative non-randomised | Q3;Q4 |
| CT0137 | low | 0 → 2 | Mixed methods → **Quantitative descriptive** | mmat_category;Q1;Q3 |
| CT0166 | low → **high** | 2 → 4 | Quantitative non-randomised | Q2;Q5 |
| CT0184 | high → **moderate** | 4 → 3 | Quantitative non-randomised → **Quantitative descriptive** | mmat_category;Q2;Q4;Q5 |
| CT0220 | low | 2 → 1 | Quantitative descriptive | Q3 |
| CT0335 | low → **high** | 2 → 4 | Mixed methods | Q3;Q4 |
| CT0348 | low → **moderate** | 2 → 3 | Quantitative descriptive | Q4;Q5 |

## 함의

이 감사는 **자동 전문 추출을 쓰는 리뷰가 품질평가에서 체계적 하향 편의를 가질 수 있음**을 보여준다. 추출 한도·OCR 실패·부록 누락은 모두 'CT'로 흘러들어가고, CT는 `n_yes`를 낮춰 등급을 떨어뜨린다. 본 리뷰는 절단이 확인된 전건을 전문으로 재평가해 해소했고, 그 과정을 프로토콜 이탈 로그에 기록한다.
