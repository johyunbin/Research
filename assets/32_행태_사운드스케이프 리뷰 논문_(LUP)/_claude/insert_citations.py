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
MS = os.path.join(os.path.dirname(BASE), "01_논문작업", "Manuscript_KO.md")

from build_numbered_refs import cr, vancouver, AUTHOR_FIX   # noqa: E402  서식 규칙 재사용

TOKEN = re.compile(r"\[\d+(?:,\d+)*\]|\{\{#?[A-Za-z0-9_]+(?:,#?[A-Za-z0-9_]+)*\}\}")


def token_keys(tok):
    """토큰 → 키 목록. `[3,4]` → ['old:3','old:4'] · `{{a,b}}` → ['a','b']
    · `{{#9,#10,a}}` → ['old:9','old:10','a'] — 기존 번호와 신규 문헌을 한 괄호로 묶을 때"""
    if tok.startswith("["):
        return [f"old:{n}" for n in tok[1:-1].split(",")]
    return [f"old:{k[1:]}" if k.startswith("#") else k for k in tok[2:-2].split(",")]

# ── 삽입 정의 ────────────────────────────────────────────────────────
# (앵커, 치환문). 앵커는 본문에 **정확히 1회** 등장해야 한다.
# 치환문에서 신규 문헌은 `{{key}}`(연속은 `{{a,b}}`), 기존 문헌 재인용은 기존 번호 `[N]`
# 을 그대로 쓴다(재번호는 이 스크립트가 알아서 한다).
#
# ★ 실행 이력 (완료된 구성은 git 이력에 있다 — 앵커가 소진되므로 재실행 불가)
#   r1 2026-08-15: Results 인용 8건 (culbert·stansfield·teo·deng·nguyen·papadakis·guski·basner)
#   r2 2026-08-16: §2.6 보고편향 문턱 근거 Sterne 2011 신설
#   r3 2026-08-16: §1.1 첫 문단 인용(사용자 메모 "Reference 추가할것") — Southworth 1969
#      "The Sonic Environment of Cities" = 도시 설계의 시각 편중·음환경 방치를 지적한 정전
#   r4 2026-09-16: 외부 AI 검토본(Manuscript_KO_20260828_FINAL) 반영 — 서론·Discussion 을
#      md 에 직접 한국어로 재작성하며 `{{key}}` 플레이스홀더를 본문에 넣었으므로 INSERTS 는
#      비우고 NEW_REFS 만 채운다. 검토본 서지 6건 중 4건 채택·메타데이터 정정, 2건 기각:
#        [44] 제1저자 Zhang Y → Zhang R · [46] Nielbo KL → FL · [49] 저널명 오류
#        (Journal of Urban Design → Journal of Planning Literature, 본문 미인용이던 것을 §1.1 에 인용)
#        기각: [48] Steele 2017 = JASA 학회 초록(1쪽) — 동일 주장을 Steele 2019 전문 논문이 뒷받침
# ⚠️ r3 앵커는 치환문이 원문을 접두로 포함해 **재실행 시 중복 삽입**된다(dry-run 실측).
#    실행 완료분은 즉시 비운다 — 다음 삽입 때 새 구성을 채울 것.
INSERTS = []

# ── 신규 배경 문헌 ───────────────────────────────────────────────────
# key → DOI. **전건 Crossref 실재 확인을 통과한 것만 여기 넣는다.** (확인일 2026-09-16)
NEW_REFS = {
    "bild2016":      "10.1177/0885412216662001",   # 계획 실무의 소리 취급 리뷰 (J Plan Lit)
    "zhang2025ctx":  "10.1121/10.0036882",         # 사운드스케이프와 맥락 프레임워크 (JASA)
    "hermida2019":   "10.3390/ijerph16040551",     # 사람–장소 상호작용 (IJERPH)
    "nielbo2013":    "10.1121/1.4800502",          # 활동 적합성·어포던스 (POMA)
    "steele2019":    "10.3390/ijerph16101865",     # Musikiosk 개입 — ★코퍼스 포함 연구(FINAL_INCLUDE)
    # r5 2026-09-17: LUP 리뷰 예시(Zhang et al. 2025)처럼 결과 절에서 메타분석 기여 연구를 개별 인용.
    #   16편 중 서지에 없던 7편(DOI = fulltext/references.csv) + 이질성 해석 기준 1편.
    "ba2020":        "10.1016/j.buildenv.2020.107104",   # uid 665 체류
    "fu2026":        "10.3390/buildings16101973",        # uid 1280 체류
    "guo2024":       "10.3390/f15101751",                # uid 1076 상관
    "zhou2026":      "10.1016/j.apacoust.2025.111101",   # uid 1221 상관
    "mansouri2025":  "10.3390/acoustics7040064",         # uid 1177 상관
    "bao2023":       "10.3390/f14122329",                # uid 980 상관
    "cao2021":       "10.1016/j.scs.2021.102860",        # uid CT0184 상관
    "higgins2003":   "10.1136/bmj.327.7414.557",         # I² 25/50/75% 해석 기준 (BMJ)
    # r6 2026-09-18: 9월 추가 전문평가로 MA4 에 편입된 연구(uid 80) — 3.4절 결과 서술에서 개별 인용
    "yu2008":        "10.1121/1.2821955",                # uid 80 상관 (JASA) · Crossref 확인 대상
}

# DOI 가 없는 1차 출처(표준·보고서·단행본) — 서지를 직접 적는다
NEW_MANUAL = {}

# Crossref 레코드의 표기 오류 보정(고유명사·중복 페이지 표기만 — 내용은 건드리지 않는다)
STRING_FIX = {
    "sterne2011": [("343:d4002-d4002", "343:d4002")],   # page 필드가 동일값 중복
    "nielbo2013": [("2013:040059-040059", "2013;19:040059")],  # Crossref 에 권호 누락(POMA 19)
}


def split_manuscript(t):
    i = t.index("## References")
    # 참고문헌 뒤 첫 부록(2026-09-17 보충자료 → 부록 A·B 전환). 구판 원고 호환을 위해 둘 다 찾는다.
    j = t.find("---\n\n## Appendix A")
    if j < 0:
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
        # 괄호 안은 오름차순(Vancouver 관례) — 번호 부여는 위에서 쓴 순서대로 이미 끝났다
        return "[" + ",".join(str(n) for n in sorted(num[k] for k in token_keys(m.group()))) + "]"

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

    # 안내 인용구는 사용자가 검토본에서 제거했다(2026-08-16) — 서지만 싣는다.
    ref = ("## References\n\n"
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
