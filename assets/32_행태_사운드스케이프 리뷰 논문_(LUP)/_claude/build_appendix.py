# -*- coding: utf-8 -*-
"""
Paper32 — 원고 부록 A·B 생성 → fulltext/appendix_ab.md

사용자 결정(2026-09-17): 보충자료 S1–S14 를 없애고 일반 리뷰 논문처럼 본문과 부록으로 마무리한다.
  Appendix A = 데이터베이스별 검색식(PRISMA 2020 item 7)
  Appendix B = 포함 연구 98편과 특성(PRISMA 2020 item 17 "cite each included study and present its characteristics")
★ 검색식은 실제로 실행한 문자열이다 — PubMed 는 문서 기재본이 아니라 `pubmed_openalex_pull.py` 실행본,
  OpenAlex 는 `openalex_pilot_search.py` 의 문자열·필터(2026-08-02 실행분 `openalex_supplement_20260802_211715.csv`).
★ 부록 B 의 연구 표기 = table1_v2.csv `study`. 저자명이 비어 있던 18편은 DOI 로 Crossref 조회한 값
  (`fulltext/appendix_b_author_lookup.json`).
"""
import os, sys, csv, json, re, importlib.util

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
sys.path.insert(0, BASE)
from make_fig7_geo_time import norm_countries   # noqa: E402


def _module_strings(path):
    """스크립트를 실행하지 않고 모듈 수준 문자열 상수(연결식 포함)만 문법 트리로 평가한다."""
    import ast
    tree = ast.parse(open(path, encoding="utf-8").read())
    env = {}

    def ev(node):
        if isinstance(node, ast.Constant) and isinstance(node.value, str):
            return node.value
        if isinstance(node, ast.Name):
            return env[node.id]
        if isinstance(node, ast.BinOp) and isinstance(node.op, ast.Add):
            return ev(node.left) + ev(node.right)
        if isinstance(node, ast.JoinedStr):
            return "".join(ev(v.value) if isinstance(v, ast.FormattedValue) else v.value
                           for v in node.values)
        raise ValueError
    for node in tree.body:
        if isinstance(node, ast.Assign) and len(node.targets) == 1 and isinstance(node.targets[0], ast.Name):
            try:
                env[node.targets[0].id] = ev(node.value)
            except (ValueError, KeyError):
                pass
    return env


def search_strings():
    """실행 스크립트에서 문자열을 그대로 읽는다(원고에 옮겨 적다 틀리지 않게)."""
    md = open(os.path.join(BASE, "search_strings_20260802_205110.md"), encoding="utf-8").read()
    blocks = re.findall(r"```[a-z]*\n(.*?)```", md, re.S)
    one = lambda s: re.sub(r"\s+", " ", s).strip()
    wos = one(next(b for b in blocks if b.lstrip().startswith("TS=")))
    scopus = one(next(b for b in blocks if b.lstrip().startswith("TITLE-ABS-KEY")))
    pubmed = one(_module_strings(os.path.join(BASE, "pubmed_openalex_pull.py"))["PUBMED_TERM"])
    oa = _module_strings(os.path.join(BASE, "openalex_pilot_search.py"))
    return {"wos": wos, "scopus": scopus, "pubmed": pubmed, "openalex": one(oa["QUERY"]),
            "oa_filters": oa["FILTERS"]}


def main():
    s = search_strings()
    missing = [k for k in ("wos", "scopus", "pubmed") if not s[k]]
    if missing:
        raise SystemExit(f"⚠️ 검색식을 읽지 못했다: {missing}")
    openalex = s["openalex"]
    assert s["oa_filters"] == "language:en,type:article,primary_location.source.type:journal", s["oa_filters"]

    # 표로 넣으면 검색식 칸이 좁아 몇 쪽으로 늘어진다 → 출처별 문단. `*`(와일드카드)는 빌더가
    # 기울임 표시로 읽지 않도록 이스케이프한다.
    esc = lambda q: q.replace("*", "\\*")
    L = ["## Appendix A. Search strategies", "",
         "All searches were run on 2 August 2026 without date restrictions. The OpenAlex search was "
         "used to identify records not indexed in the three bibliographic databases.", "",
         "Web of Science Core Collection (limits: English; document type, article):", "",
         esc(s["wos"]), "",
         "Scopus (limits included in the string):", "",
         esc(s["scopus"]), "",
         "PubMed (limits included in the string):", "",
         esc(s["pubmed"]), "",
         "OpenAlex, title and abstract search (filters: English; type, article; source type, journal):", "",
         esc(openalex), ""]

    look = json.load(open(os.path.join(FT, "appendix_b_author_lookup.json"), encoding="utf-8"))
    rows = [r for r in csv.DictReader(open(os.path.join(FT, "table1_v2.csv"), encoding="utf-8-sig"))
            if r["verdict"] == "FINAL_INCLUDE"]
    # ★ 2026-09-18 추가 전문평가 반영: 기대 편수 리터럴(98) 대신 판정 정본(corpus_v4_verdicts.csv)의
    #   FINAL_INCLUDE 수와 대조한다 — Table 1 이 판정 정본과 어긋나면 멈추는 검사 목적은 그대로다.
    n_inc = sum(1 for v in csv.DictReader(open(os.path.join(FT, "corpus_v4_verdicts.csv"), encoding="utf-8-sig"))
                if v["final_verdict"] == "FINAL_INCLUDE")
    if len(rows) != n_inc:
        raise SystemExit(f"⚠️ Table 1 포함 연구 수({len(rows)})가 판정 정본 FINAL_INCLUDE({n_inc})와 다르다")
    ext_country = {e["uid"]: e["country"] for e in
                   csv.DictReader(open(os.path.join(FT, "corpus_v4_extraction.csv"), encoding="utf-8-sig"))}
    # ★ 2026-09-18 게이트 지적: 본문에 번호로 인용된 연구는 참고문헌(권호 연도)과 연도를 맞춘다.
    #   코퍼스 연도는 온라인 게재 연도라 CT0126(2022 → 2023)·941(2023 → 2024)이 참고문헌과 달랐다.
    uid_doi = {u: v.get("doi", "") for u, v in look.items() if isinstance(v, dict)}
    uid_doi.update({r["uid"]: r["doi"] for r in csv.DictReader(open(os.path.join(FT, "references.csv"), encoding="utf-8-sig"))
                    if r.get("doi")})
    ref_year = {}
    for m in re.finditer(r"^\[\d+\] (.+)$", open(os.path.join(FT, "references_numbered.md"), encoding="utf-8").read(), re.M):
        d, y = re.search(r"doi:(\S+)", m.group(1)), re.findall(r"\. (\d{4});", m.group(1))
        if d and y:
            ref_year[d.group(1).lower().rstrip(".")] = y[-1]
    out = []
    for r in rows:
        if r["study"].startswith("["):
            lk = look.get(r["uid"])
            if not lk or "label" not in lk:
                raise SystemExit(f"⚠️ 저자명 조회값 없음: {r['uid']}")
            lab = lk["label"]
            lab = lab if not lab.isupper() else lab.title()
            # ★ 2026-09-18: 연도는 코퍼스(추출표) 연도로 통일 — Crossref issued 연도와 달라 Table 1·본문과 어긋나던 문제
            study = f"{lab} ({r['year'] or lk['year']})"
        else:
            study = r["study"]
        ry = ref_year.get((uid_doi.get(r["uid"]) or "").lower())
        if ry:
            study = re.sub(r"\((\d{4})\)$", f"({ry})", study)
        cap = lambda v: "NR" if v in ("", "NR") else v[0].upper() + v[1:]
        # Fig. 7·Table 1 과 같은 원천(추출표 country)·같은 정규화 — 국가별 수가 정확히 일치한다
        country = "; ".join(norm_countries(ext_country[r["uid"]])) or "NR"
        setting = {"lab(outdoor scene)": "laboratory (outdoor scene)",
                   "recreation": "recreation area"}.get(r["setting"], r["setting"])
        design = {"lab experiment": "laboratory experiment",
                  "observational": "observation"}.get(r["design"], r["design"])   # Table 1·2 와 같은 말
        out.append((study, country, cap(setting), cap(design),
                    cap(r["behaviour_domain"].replace("space-use", "space use")),
                    cap(r["direction"]), cap(r["quality"])))
    out.sort(key=lambda x: (x[0].lower(), x[0]))
    # 사용자 ver6(2026-09-17): 부록 B 는 가로 쪽. 표시는 build_docx.py 가 구역 나눔으로 바꾼다.
    L += ["<<LANDSCAPE>>", "", "## Appendix B. Characteristics of included studies", "",
          # 표 제목은 짧게, 정의는 표 아래 주석으로(LUP 예시 Zhang et al. 2025 의 표 구성, 2026-09-17)
          f"**Table B1.** Studies included in the review (n = {len(out)}).", "",
          "| Study | Country | Setting | Design | Behavioural domain | Direction | MMAT |",
          "|---|---|---|---|---|---|---|"]
    L += [f"| {' | '.join(c.replace('|', '/') for c in row)} |" for row in out]
    L += ["", "*Note.* A study can contribute to more than one behavioural domain. Direction: forward = "
          "acoustic environment to behaviour; reverse = behaviour or activity to acoustic environment "
          "or soundscape; both = both directions examined. MMAT = Mixed Methods Appraisal Tool 2018 "
          "grade. NR = not reported."]
    text = "\n".join(L) + "\n"
    open(os.path.join(FT, "appendix_ab.md"), "w", encoding="utf-8").write(text)
    print(f"[저장] fulltext/appendix_ab.md · 부록 A 4개 검색식 · 부록 B {len(out)}편")


if __name__ == "__main__":
    main()
