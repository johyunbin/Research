# -*- coding: utf-8 -*-
"""
Paper32 — 번호 인용 원고에 인용을 **삽입하고 전면 재번호**한다.

`build_numbered_refs.py` 는 저자-연도 → `[N]` 일회성 변환기라 이미 번호가 된 원고에는
재실행할 수 없다(CITE 딕셔너리의 원본 문자열이 본문에 남아 있지 않다). 그런데 번호는
**본문 첫 등장 순서**(Vancouver)이므로 §3.2 에 인용 하나를 넣으면 그 뒤 번호가 전부 밀린다.
번호를 손으로 옮기는 것은 사고의 원인이므로 이 스크립트가 대신한다.

동작
  1. 원고를 body / References / Supplementary 로 분리하고 References 를 {N: 서지} 로 읽는다.
     **서지의 단일 소스는 원고 자신**이다(별도 상태 파일을 만들지 않는다).
  2. INSERTS 의 앵커를 본문에서 찾아 치환한다. 앵커가 정확히 1회 등장하지 않으면 중단한다.
  3. 본문을 위치 순서로 훑어 `[N]`·`{{key}}` 토큰을 모으고 첫 등장 순으로 재번호한다.
  4. 본문 토큰을 일괄 재작성하고 References 블록을 다시 만든다.

서지 서식(Vancouver·paper31 실측)은 `build_numbered_refs.py` 의 함수를 그대로 쓴다.

사용
  python insert_citations.py --dry-run   # 원고를 건드리지 않고 결과만 출력
  python insert_citations.py             # 원고 in-place + fulltext/references_numbered.md
"""
import sys, os, re, time, argparse

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
FT = os.path.join(BASE, "fulltext")
MS = os.path.join(os.path.dirname(BASE), "01_논문작업", "Manuscript_KO_20260806.md")

from build_numbered_refs import cr, vancouver, AUTHOR_FIX   # noqa: E402  서식 규칙 재사용

TOKEN = re.compile(r"\[\d+(?:,\d+)*\]|\{\{[A-Za-z0-9_]+(?:,[A-Za-z0-9_]+)*\}\}")


def token_keys(tok):
    """토큰 → 키 목록. `[3,4]` → ['old:3','old:4'] · `{{a,b}}` → ['a','b']"""
    if tok.startswith("["):
        return [f"old:{n}" for n in tok[1:-1].split(",")]
    return tok[2:-2].split(",")

# ── 삽입 정의 ────────────────────────────────────────────────────────
# (앵커, 치환문). 앵커는 본문에 **정확히 1회** 등장해야 한다.
# 치환문에서 신규 문헌은 `{{key}}`(연속은 `{{a,b}}`), 기존 문헌 재인용은 기존 번호 `[N]`
# 을 그대로 쓴다(재번호는 이 스크립트가 알아서 한다).
#
# ★ 문장 표현이 함께 바뀐 곳이 있다. 인용을 붙이려다 보니 **원문이 인용 가능한 범위를
#   넘어서는 주장을 하고 있다는 것이 드러났기 때문**이다(Codex 독립 판정, 2026-08-15).
#   - §3.1 "WoS·Scopus·PubMed에 없다 = 누락 근거 약함" → 세 DB 전체의 불완전성은 후보
#     문헌이 입증하지 않는다. 이 리뷰가 실제로 보인 것(밖에서 2편 추가)으로 범위를 좁혔다.
#   - §3.1 "제목만으로 판단 = 사실상 무작위" → Mateen(2013)은 오히려 반증에 가깝고
#     (제목-only 도 recall 100%), 내부 0/17 로도 "무작위"는 입증되지 않는다. 표현을 낮췄다.
#   - §3.1 "초록 확보 가능성이 정확도를 결정했다" → 두 집단이 무작위 배정이 아니므로 인과
#     표현을 연관 표현으로 바꿨다.
#   - §3.2 "문화적 조건화는 이 분야의 기본 전제" → 후보 문헌은 차이의 보고를 지지하지만
#     "분야의 기본 전제"라는 지위까지는 입증하지 않는다.
#   - §3.7 "수백 편의 연구가 있다" → 어떤 후보도 이 수량을 입증하지 못한다. 검증 가능한
#     진술로 바꾸고 gap 주장은 코퍼스 범위로 한정했다.
INSERTS = [
    # §3.1 OpenAlex 색인 범위
    ("OpenAlex가 세 데이터베이스보다 훨씬 넓게 색인하기 때문이며, 따라서 "
     "**WoS·Scopus·PubMed에 없다는 사실은 누락의 근거로 약하다.** 색인 범위가 넓다는 것은 "
     "관련 문헌을 더 많이 담는다는 뜻이기도 하지만, 걸러야 할 것도 그만큼 많이 담는다는 뜻이다.",
     "OpenAlex의 색인이 Web of Science·Scopus보다 훨씬 크기 때문이다{{culbert2025}}. "
     "체계적 문헌고찰에서 OpenAlex를 검색원으로 쓰면 관행적 검색원보다 훨씬 많은 레코드가 "
     "나오고 그만큼 스크리닝 부담이 커진다는 것도 이미 보고돼 있다{{stansfield2025}}. "
     "색인 범위가 넓다는 것은 관련 문헌을 더 많이 담는다는 뜻이기도 하지만, 걸러야 할 것도 "
     "그만큼 많이 담는다는 뜻이다. 세 데이터베이스 밖에서 2편이 추가됐다는 사실 자체는 "
     "**데이터베이스 검색만으로 코퍼스가 닫히지 않았음**을 보여 준다."),

    # §3.1 제목만으로 하는 스크리닝
    ("**초록 확보 가능성이 스크리닝 정확도를 결정했다.** 이 관찰은 두 방향으로 쓰인다. "
     "하나는 인용추적처럼 초록 확보가 어려운 경로에서 제목만으로 판단하는 것이 사실상 "
     "무작위에 가깝다는 경고이고, 다른 하나는 미확보로 남은 제목-판단 레코드 22건이 "
     "코퍼스를 바꿨을 가능성이 낮다는 근거다.",
     "**초록 확보 여부가 최종 포함과 강하게 연관됐다.** 제목만 보는 1차 선별이 제목과 초록을 "
     "함께 보는 선별보다 민감도가 낮다는 것은 방법론 연구에서도 보고된 바 있다{{teo2023}}. "
     "다만 초록을 확보한 레코드와 그러지 못한 레코드는 무작위로 나뉜 집단이 아니고 후자의 "
     "표본도 17건에 지나지 않으므로, 미확보로 남은 제목-판단 레코드 22건이 코퍼스에 미쳤을 "
     "영향은 배제할 수 없다."),

    # §3.2 사운드스케이프 평가의 문화 조건화
    ("사운드스케이프 평가가 문화적으로 조건화된다는 것은 이 분야의 기본 전제 중 하나이므로, "
     "합성 추정치를 문화적으로 일반적인 값으로 읽어서는 안 된다.",
     "교차문화 연구는 같은 유형의 도시 옥외공간이라도 국가에 따라 사운드스케이프 평가가 "
     "달라진다고 보고해 왔다{{deng2020,nguyen2026}}. 이런 비교는 지각 속성(perceptual "
     "attributes)의 번역과 문화 간 적응(cross-cultural adaptation) 절차에 따라 결과가 "
     "달라질 수 있어 그 자체로 까다롭기도 하다{{papadakis2022}}. 어느 쪽이든 합성 추정치를 "
     "문화적으로 일반적인 값으로 읽어서는 안 된다."),

    # §3.7 항공기 소음 문헌의 규모
    ("항공기 소음은 환경소음 연구에서 가장 방대한 문헌을 가진 주제 중 하나이며 건강 영향과 "
     "짜증도에 관해서는 수백 편의 연구가 있다. 그런데 **공항 주변 사람들이 옥외공간에서 "
     "무엇을 하는지는 거의 측정된 적이 없다.**",
     "항공기 소음은 환경소음 연구에서 축적이 두터운 주제다. 세계보건기구 유럽지역 환경소음 "
     "지침을 뒷받침한 체계적 문헌고찰만 보아도 짜증도(annoyance){{guski2017}}와 "
     "수면{{basner2018}}에서 항공기 소음은 별도의 종합 대상이 될 만큼 연구가 모여 있고, "
     "건강 영향 전반에서도 사정은 다르지 않다[4]. 그런데 **공항 주변 옥외공간에서 사람들이 "
     "무엇을 하는지를 측정한 연구는 우리 적격성 기준 안에서 그 2편이 전부였다.**"),
]

# ── 신규 배경 문헌 ───────────────────────────────────────────────────
# key → DOI. **전건 Crossref 실재 확인을 통과한 것만 여기 넣는다.**
# 확인일 2026-08-15 · 8/8건 DOI 조회 성공 · 제목·연도·저널 대조 완료
NEW_REFS = {
    "culbert2025":    "10.1007/s11192-025-05293-3",   # OpenAlex vs WoS·Scopus 코퍼스 규모
    "stansfield2025": "10.1002/cesm.70038",           # OpenAlex 검색량·스크리닝 부담(Cochrane ESM)
    "teo2023":        "10.1186/s13643-023-02374-3",   # title-only vs title+abstract 민감도
    "deng2020":       "10.3390/app10030960",          # 중국·크로아티아 옥외공간 교차국가 비교
    "nguyen2026":     "10.1016/j.apacoust.2026.111414",  # 프랑스·일본·베트남 참가자 비교
    "papadakis2022":  "10.1016/j.apacoust.2022.109031",  # 지각 속성 번역·문화 간 적응 방법론
    "guski2017":      "10.3390/ijerph14121539",       # WHO 지침 근거 — 환경소음과 annoyance
    "basner2018":     "10.3390/ijerph15030519",       # WHO 지침 근거 — 환경소음과 수면
}

# DOI 가 없는 1차 출처(표준·보고서·단행본) — 서지를 직접 적는다
NEW_MANUAL = {}

# Crossref 레코드의 표기 오류 보정(고유명사 대소문자만 — 내용은 건드리지 않는다)
STRING_FIX = {
    "stansfield2025": [("Openalex", "OpenAlex")],   # 출판사가 title-case 로 기탁하며 뭉갬
}


def split_manuscript(t):
    i = t.index("## References")
    j = t.index("---\n\n## Supplementary material")
    return t[:i], t[i:j], t[j:]


def parse_refs(block):
    """References 블록 → {번호: 서지문자열}"""
    refs = {}
    for m in re.finditer(r"^\[(\d+)\]\s+(.+?)(?=\n\n\[|\n\n---|\Z)", block, re.M | re.S):
        refs[int(m.group(1))] = " ".join(m.group(2).split())
    if not refs:
        raise SystemExit("⚠️ References 를 한 건도 파싱하지 못했다 — 서식이 바뀌었는지 확인할 것")
    miss = [n for n in range(1, max(refs) + 1) if n not in refs]
    if miss:
        raise SystemExit(f"⚠️ References 번호가 연속이 아니다: 빠진 번호 {miss}")
    return refs


def apply_inserts(body):
    for anchor, repl in INSERTS:
        c = body.count(anchor)
        if c != 1:
            raise SystemExit(f"⚠️ 앵커가 {c}회 등장한다(1회여야 한다): {anchor[:70]!r}")
        body = body.replace(anchor, repl)
    return body


def renumber(body, refs):
    """첫 등장 순으로 재번호. 반환: (새 body, 순서대로의 키 목록, old→new 대응)"""
    order, num = [], {}
    for m in TOKEN.finditer(body):
        tok = m.group()
        for k in token_keys(tok):
            if k.startswith("old:") and int(k[4:]) not in refs:
                raise SystemExit(f"⚠️ 본문 {tok} 이 References 에 없다")
            if k not in num:
                num[k] = len(order) + 1
                order.append(k)

    def sub(m):
        return "[" + ",".join(str(num[k]) for k in token_keys(m.group())) + "]"

    return TOKEN.sub(sub, body), order, num


def build_bib(order, refs):
    """순서대로의 키 목록 → [(번호, 서지)] · 신규 키는 Crossref 조회"""
    out, fetched = [], []
    for i, k in enumerate(order, 1):
        if k.startswith("old:"):
            out.append((i, refs[int(k[4:])]))
            continue
        if k in NEW_MANUAL:
            out.append((i, NEW_MANUAL[k])); fetched.append((k, "(직접 표기)", NEW_MANUAL[k]))
            continue
        doi = NEW_REFS.get(k)
        if not doi:
            raise SystemExit(f"⚠️ 신규 인용 키 '{k}' 의 DOI 가 NEW_REFS 에 없다")
        m = cr(doi)
        if not m:
            raise SystemExit(f"⚠️ Crossref 조회 실패: {k} (doi:{doi}) — 확인 못 한 문헌은 인용하지 않는다")
        v = vancouver(m)
        if k in AUTHOR_FIX:
            v = AUTHOR_FIX[k] + ". " + v.lstrip(". ")
        for a, b in STRING_FIX.get(k, []):
            v = v.replace(a, b)
        out.append((i, v)); fetched.append((k, doi, v))
        time.sleep(0.35)
    return out, fetched


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    a = ap.parse_args()

    t = open(MS, encoding="utf-8").read()
    body, refblock, tail = split_manuscript(t)
    refs = parse_refs(refblock)
    print(f"[읽음] 본문 인용 토큰 {len(TOKEN.findall(body))}개소 · 서지 {len(refs)}건")

    body = apply_inserts(body)
    body, order, num = renumber(body, refs)
    bib, fetched = build_bib(order, refs)

    # ── 검증 ────────────────────────────────────────────────────────
    used = set()
    for tok in re.findall(r"\[\d+(?:,\d+)*\]", body):
        used |= {int(x) for x in tok[1:-1].split(",")}
    expect = set(range(1, len(order) + 1))
    if used != expect:
        raise SystemExit(f"⚠️ 번호 무결성 실패 — 본문 미등장 {sorted(expect - used)} · "
                         f"서지 없음 {sorted(used - expect)}")
    if len(bib) != len(order):
        raise SystemExit("⚠️ 서지 건수와 인용 건수가 다르다")

    ref = ("## References\n\n"
           "> 번호는 본문 첫 등장 순서다. 포함 98편의 전건 목록은 Supplementary S5, "
           "전문 단계 배제 문헌과 사유는 S6에 있다. 배경 문헌은 모두 Crossref에서 서지를 "
           "확인했다.\n\n"
           + "\n\n".join(f"[{n}] {s}" for n, s in bib) + "\n\n")

    # ── 보고 ────────────────────────────────────────────────────────
    moved = [(int(k[4:]), v) for k, v in num.items() if k.startswith("old:") and int(k[4:]) != v]
    print(f"[재번호] 총 {len(order)}건 · 번호가 바뀐 기존 문헌 {len(moved)}건")
    for o, n in sorted(moved):
        print(f"         [{o}] → [{n}]  {refs[o][:64]}")
    if fetched:
        print(f"[신규 문헌] {len(fetched)}건")
        for k, doi, v in fetched:
            print(f"         {k:22s} {doi}\n           {v}")

    if a.dry_run:
        print("\n[dry-run] 원고를 쓰지 않았다.")
        return
    open(MS, "w", encoding="utf-8").write(body + ref + tail)
    open(os.path.join(FT, "references_numbered.md"), "w", encoding="utf-8").write(
        "\n\n".join(f"[{n}] {s}" for n, s in bib) + "\n")
    print(f"\n[저장] 원고 in-place · fulltext/references_numbered.md")


if __name__ == "__main__":
    main()
