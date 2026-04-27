# archive/

중간 산출물 보관소. 최종본은 모두 상위 `ETC/`로 옮겨졌고 여기는 작업 과정 기록만 남는다.

## 구조

```
archive/
├── drafts/       Portfolio_HyunInJo.pptx, Portfolio_HyunInJo_v2.pptx
│                 → 스크립트로 자동 생성한 1, 2차 초안
│                 → 최종본은 ETC/Portfolio_HyunInJo_Final.pptx (수동 마무리)
│
├── pdfs/         slide_versions_full.pdf, slide_content_versions.pdf
│                 → generate_slide_*.py가 만든 슬라이드 비교용 PDF
│
├── exploration/  init/recom/design_options/slide_content_versions.docx
│                 design_samples.html, s6_redesign.html
│                 → 디자인 브레인스토밍·시안 단계 산출물
│
└── scripts/      build_presentation.py, build_presentation_v2.py
                  generate_slide_pdf.py, generate_slide_content_pdf.py
                  package.json, package-lock.json
                  └─ lib/  기존 scripts/ 라이브러리 (pptx_helpers.py 등)
                  → 일회성 생성 스크립트 + Python/Node 의존성
```

## 재실행 가이드

루트 readme 참고. 스크립트 재실행이 필요하면:

```bash
cd archive/scripts
python3 -m venv .venv && source .venv/bin/activate
pip install python-pptx fpdf2
# Node 의존성 (docx 라이브러리)이 필요하면:
npm install
```

`lib/` 모듈은 `archive/scripts/build_presentation*.py` 가 `from scripts import ...` 형태로 import 했었으므로,
재실행 시 import path 수정이 필요할 수 있다.
