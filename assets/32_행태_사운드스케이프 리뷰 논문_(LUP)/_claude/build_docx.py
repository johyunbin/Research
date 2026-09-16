# -*- coding: utf-8 -*-
"""
Paper32 — 원고 markdown → Word(.docx)

★ 서식 정본 = `templates/Manuscript_KO_format_20260917.docx`
  사용자가 빌드본 ver3 을 Word 에서 직접 고친 파일이다(2026-09-17 "이 파일을 기반으로 향후
  업데이트"). 스타일·테마·페이지·문서 설정은 이 파일을 그대로 쓰고, 문단은 역할마다 이 파일
  안의 **견본 문단에서 pPr·rPr 를 복제**해 만든다. 서식 값을 이 스크립트에 옮겨 적지 않는다 —
  사용자가 서식을 다시 고치면 템플릿 파일만 바꾸면 되고, 옮겨 적다 생기는 누락도 없다.
  (구판은 paper31 실측값을 코드에 박아 두어, 사용자가 고칠 때마다 역산해야 했다.)
  원고 문장은 템플릿이 아니라 md 에서 온다. 템플릿의 본문은 견본을 읽은 뒤 비운다.
출력: 01_논문작업/Manuscript_{KO|EN}_{YYYYMMDD}_ver{N}.docx (구판은 자동으로 old/ 이관)
"""
import sys, os, re, datetime
from copy import deepcopy

from docx import Document
from docx.shared import Inches
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(BASE)
SRC = os.path.join(ROOT, "01_논문작업",
                   sys.argv[1] if len(sys.argv) > 1 else "Manuscript_KO.md")
FIGDIR = os.path.join(BASE, "figures")
TEMPLATE = os.path.join(BASE, "templates", "Manuscript_KO_format_20260917_ver1.docx")
#   ver1(2026-09-17 서론 점검본): 본문 양쪽 정렬 · 절/소절 제목 앞 빈 줄 — 사용자가 서론에서 직접 고친 서식

# 그림별 삽입 폭(in) — ★ 삽입 폭 = 작화 폭(viz_theme.W_FULL = 6.05 in), 축소 없이 1:1.
#   그림을 크게 그려 놓고 줄여 넣으면 글자가 4~6 pt 로 떨어진다(구판 실측).
FIG_W = {}
_FIG_DEFAULT_W = 6.05

# 그림 캡션 — ★숫자는 figures/fig_data.json 에서 읽는다(캡션도 하드코딩하지 않는다).
import json as _json
_FD = _json.load(open(os.path.join(FIGDIR, "fig_data.json"), encoding="utf-8"))
_MA = _FD["ma"]
_G = _FD["generations"]
_Q = _FD["quality"]["items"]


def _pl(k):
    return _MA[k]["pooled"]


CAPTIONS = {
    "Fig1_PRISMA": (
        "PRISMA 2020 flow diagram covering all three registered identification routes: database "
        "searching, citation tracking of every included study, and a supplementary index. "
        f"The dominant exclusion reason in the database branch is animal or wildlife research "
        f"({_FD['prisma']['db']['excl'][0][1]} records), reflecting the terminological overlap "
        "between soundscape ecology and human soundscape research. In the supplementary branch, "
        f"{_FD['prisma']['supp']['excl'][0][1]} of {_FD['prisma']['supp']['excluded']} exclusions "
        "were non-empirical items such as book reviews and editorials, which is why absence from "
        "the three databases is weak evidence that a study was missed."),
    "Fig2_Forest": (
        "Pooled estimates for four behavioural clusters, from random-effects REML with the "
        "Hartung–Knapp adjustment. Squares are individual effects, sized by their "
        "random-effects weight (printed with each effect and its 95% CI in the right-hand "
        "columns); diamonds are pooled estimates; triangles mark studies retrieved by citation "
        "searching. Panel (d) is displayed on the r scale. "
        f"Only social interaction (p = {_pl('social')['p']:.3f}) and the sound–behaviour "
        f"correlation (p = {_pl('correlation')['p']:.3f}) exclude zero, and both do so for the "
        "mean effect only — the 95% prediction interval includes zero in all four clusters."),
    "Fig3_EvidenceMap": (
        "Evidence map of behavioural domain by sound source. Cells count studies, and a study "
        "contributes to every cell it covers. Every combination is populated except in the "
        "aircraft-noise column, where three of the five domains are empty and the remaining "
        f"cells hold {_FD['n_aircraft_records']} domain-level records from "
        f"{_FD['n_aircraft_studies']} studies — the clearest gap given the size of the "
        "aircraft-noise health literature."),
    "Fig4_Direction": (
        "Direction of the studied relationship, by behavioural domain. The unit is the "
        f"study × behavioural-domain record. {_FD['direction']['total_reverse']} of "
        f"{_FD['direction']['n_directional']} directional records "
        f"({_FD['direction']['pct_reverse_of_directional']}%) run in reverse. Movement and "
        "staying are dominated by forward designs, whereas space use, activity and social "
        "behaviour approach parity between the two directions — the empirical basis for a "
        "bidirectional framing."),
    "Fig5_Methods": (
        "Behavioural measurement methods over time. Sensing, GPS, video and big-data measurement "
        f"grew from {_G['G3'][1]} studies in 2010–2019 to {_G['G3'][2]} from 2020, while "
        f"self-report grew from {_G['G1'][1]} to {_G['G1'][2]} and systematic observation from "
        f"{_G['G2'][1]} to {_G['G2'][2]}. Generations accumulate rather than replace one another; "
        f"{_G['multi_generation_studies']} studies use two or more concurrently and therefore "
        "appear in more than one series."),
    # ★ 2026-09-16 외부 AI 검토본 캡션 채택 — 정량 정보는 개념도에서 뺐으므로 캡션도 개념만.
    "Fig6_Framework": (
        "Reciprocal evidence framework linking context, acoustic environment, soundscape "
        "appraisal, and observable behaviour. Spatial, physical, and socio-cultural context "
        "shapes the acoustic environment and moderates how acoustic conditions are interpreted "
        "and acted upon. Acoustic conditions may influence observable behaviour directly or "
        "through soundscape appraisal. Behaviour and activity can, in turn, modify the acoustic "
        "environment through occupancy and human sound production. The two pathways represent a "
        "reciprocal evidence structure rather than a demonstrated closed causal feedback loop. "
        "Behavioural outcomes are organised below the framework along an engagement gradient "
        "from avoidance and passing to staying, interacting, and appropriating; this gradient is "
        "used as a synthesis device and is not a validated behavioural scale. The figure is "
        "conceptual; quantitative effect sizes, p-values, and study-quality information are "
        "reported separately in the Results and evidence tables."),
    "Fig7_GeoTime": (
        f"Geographic and temporal distribution of the {_FD['n_included']} included studies. "
        f"(a) {_FD['geo']['countries'][0][1]} studies "
        f"({_FD['geo']['countries'][0][1] / _FD['n_included'] * 100:.0f}%) were conducted in "
        f"{_FD['geo']['countries'][0][0]}. (b) {_FD['n_since_2020']} studies "
        f"({_FD['n_since_2020'] / _FD['n_included'] * 100:.0f}%) appeared in 2020 or later, and "
        "all but two reverse-direction studies appeared from 2016 onwards, so the bidirectional "
        "evidence base is younger still than the corpus as a whole. Multi-country studies are "
        "counted once per country; two studies did not report a country."),
    "Fig8_Quality": (
        "MMAT 2018 appraisal. (a) Grade distribution within each MMAT category. (b) Selected "
        "items, grouped to show the pattern that drives the Discussion: what the studies report "
        "well concerns the measurement, and what they do not report concerns the people. Sample "
        f"representativeness was clearly met in {_Q['4.2']['Y']} of {_Q['4.2']['n']} quantitative "
        f"descriptive studies, low non-response bias in {_Q['4.4']['Y']} of {_Q['4.4']['n']}, and "
        f"control of confounding in {_Q['3.4']['Y']} of {_Q['3.4']['n']} non-randomised studies. "
        "The two randomised studies report neither the randomisation procedure, nor baseline "
        "comparability, nor blinding, so their design cannot be credited at all. Grey means the "
        "information is absent, not that the study is known to be biased. The full 25-item set is "
        "Supplementary S10."),
}


# ═══════════════════ 서식 견본 ═══════════════════
_W14 = "{http://schemas.microsoft.com/office/word/2010/wordml}"
_NOISE = {qn("w:lastRenderedPageBreak"), qn("w:proofErr"), qn("w:bookmarkStart"),
          qn("w:bookmarkEnd"), qn("w:commentRangeStart"), qn("w:commentRangeEnd")}


def _clean(el):
    """견본 복제본에서 편집 흔적(rsid·w14 식별자·렌더링 표지)을 지운다."""
    el = deepcopy(el)
    for x in list(el.iter()):
        if x.tag in _NOISE and x.getparent() is not None:
            x.getparent().remove(x)
            continue
        for a in list(x.attrib):
            if "rsid" in a or a.startswith(_W14):
                del x.attrib[a]
    return el


def _base_rpr(p):
    """문단의 대표 런 서식 = 글자가 가장 긴 런의 rPr(라벨처럼 짧은 강조 런을 피한다)."""
    runs = [r for r in p.runs if r.text]
    if not runs:
        return None
    rpr = max(runs, key=lambda r: len(r.text))._r.rPr
    return _clean(rpr) if rpr is not None else None


class Fmt:
    """템플릿 문단을 역할별 견본으로 읽어 둔다. 견본을 못 찾으면 조용히 넘어가지 않고 멈춘다."""

    def __init__(self, tpl):
        P = tpl.paragraphs
        T = [p.text.strip() for p in P]

        def at(pred, what):
            for i, s in enumerate(T):
                if pred(s):
                    return i
            raise SystemExit(f"⚠️ 서식 템플릿에서 '{what}' 견본 문단을 찾지 못했다: {TEMPLATE}")

        i_abs = at(lambda s: s == "ABSTRACT", "ABSTRACT 제목")
        i_h1 = at(lambda s: s == "1. Introduction", "1수준 제목")
        i_h2 = at(lambda s: s.startswith("1.1 "), "2수준 제목")
        i_decl = at(lambda s: s == "CRediT authorship contribution statement", "선언부 제목")
        i_supp = at(lambda s: s == "Supplementary material", "보충자료 제목")
        i_corr = at(lambda s: s.startswith("*Send correspondence"), "교신저자")
        i_h1b = at(lambda s: s == "2. Methods", "2수준 절 제목")
        role = {
            "title": 0, "author": 1, "affil": 2, "fm_gap": 3,
            "corr_label": at(lambda s: s == "[Corresponding author]", "교신저자 라벨"),
            "corr": i_corr, "address": i_corr + 1,
            "abs_head": i_abs, "abs_body": i_abs + 1,
            "keywords": at(lambda s: s.startswith("Keywords"), "키워드"),
            "h1": i_h1, "h2": i_h2,
            # 본문 견본 = 1.1 아래 양쪽 정렬 문단(사용자가 서론에서 정렬을 바꿨다 — 첫 문단만 보면 놓친다)
            "body": next((k for k in range(i_h2 + 1, i_h1b) if T[k] and P[k]._p.pPr is not None
                          and P[k]._p.pPr.find(qn("w:jc")) is not None), None),
            "head_gap": at(lambda s: s.startswith("1.3 "), "1.3 소절 제목") - 1,
            "rq": at(lambda s: s.startswith("RQ1."), "연구 질문 목록"),   # 글머리 기호 목록(사용자 2026-09-17)
            "fig_cap": at(lambda s: s.startswith("Fig. 1."), "그림 캡션"),
            "tbl_cap": at(lambda s: s.startswith("Table 1."), "표 캡션"),
            "decl_body": i_decl + 1, "decl_gap": i_decl + 2,
            "ref": at(lambda s: s.startswith("[1] "), "참고문헌"),
        }
        if role["body"] is None:
            raise SystemExit("⚠️ 서식 템플릿: 서론에서 양쪽 정렬된 본문 견본을 찾지 못했다")
        if T[role["head_gap"]]:
            raise SystemExit("⚠️ 서식 템플릿: 1.3 제목 앞 문단이 빈 줄이 아니다")
        self.pPr, self.rPr = {}, {}
        for k, i in role.items():
            pp = P[i]._p.pPr
            self.pPr[k] = _clean(pp) if pp is not None else None
            self.rPr[k] = _base_rpr(P[i])
        # 쪽 나눔 문단은 자리마다 견본이 조금씩 달라(사용자 편집 흔적) 네 곳을 각각 통째로 보관
        self.brk = {}
        for k, i in {"abstract": i_abs - 1, "intro": i_h1 - 1,
                     "decl": i_decl - 1, "supp": i_supp - 1}.items():
            if not P[i]._p.findall(".//" + qn("w:br")):
                raise SystemExit(f"⚠️ 서식 템플릿: '{T[i + 1]}' 앞 문단이 쪽 나눔이 아니다")
            self.brk[k] = _clean(P[i]._p)
        pic = next(p for p in P if p._p.findall(".//" + qn("w:drawing")))
        self.pPr["fig_pic"] = _clean(pic._p.pPr)
        tbl = tpl.tables[0]
        self.tbl_after = _clean(tbl._tbl.getnext())
        cell = tbl.rows[0].cells[0].paragraphs[0]
        self.cell_pPr = _clean(cell._p.pPr)
        self.cell_head = _base_rpr(cell)
        body_run = next(r for row in tbl.rows[1:] for c in row.cells for q in c.paragraphs
                        for r in q.runs if r.text and not r.font.bold)
        self.cell_body = _clean(body_run._r.rPr)


def _append_clone(doc, el):
    doc.element.body.find(qn("w:sectPr")).addprevious(deepcopy(el))


def _set_ppr(p, pPr):
    old = p._p.pPr
    if old is not None:
        p._p.remove(old)
    if pPr is not None:
        p._p.insert(0, deepcopy(pPr))


INLINE = re.compile(r"(\*\*.+?\*\*|(?<!\*)\*[^*\n]+?\*(?!\*)|`[^`]+?`)", re.S)
SUPER = re.compile(r"([¹²³⁴⁵⁶⁷⁸⁹⁰]+)")
_SUP_DIGIT = str.maketrans("¹²³⁴⁵⁶⁷⁸⁹⁰", "1234567890")


def add_run(p, text, base, bold=False, italic=False, sup=False):
    """견본 rPr 를 복제한 런. sup=True(저자·소속 줄)면 유니코드 위첨자 숫자를 위첨자 서식의
    일반 숫자로 바꾼다 — 본문의 τ²·I² 는 사용자 서식본에서도 유니코드 그대로다."""
    for piece in (SUPER.split(text) if sup else [text]):
        if not piece:
            continue
        r = p.add_run(piece.translate(_SUP_DIGIT) if sup else piece)
        if base is not None:
            r._r.insert(0, deepcopy(base))
        if bold:
            r.font.bold = True
        if italic:
            r.font.italic = True
        if sup and SUPER.fullmatch(piece):
            r.font.superscript = True


def add_rich(p, text, base, sup=False):
    """**bold** · *italic* · `code` 를 견본 런 서식 위에 얹는다"""
    text = text.replace("\\*", "\u0001")
    for tok in INLINE.split(text):
        if not tok:
            continue
        tok = tok.replace("\u0001", "*")
        if tok.startswith("**") and tok.endswith("**") and len(tok) > 4:
            add_run(p, tok[2:-2], base, bold=True, sup=sup)
        elif tok.startswith("*") and tok.endswith("*") and len(tok) > 2:
            add_run(p, tok[1:-1], base, italic=True, sup=sup)
        elif tok.startswith("`") and tok.endswith("`") and len(tok) > 2:
            add_run(p, tok[1:-1], base, sup=sup)
        else:
            add_run(p, tok, base, sup=sup)


def _no_gap_needed(doc):
    """본문 끝이 빈 문단(표 뒤 빈 줄·쪽 나눔)이거나 제목(outlineLvl)이면 제목 앞 빈 줄을 넣지 않는다."""
    last = doc.element.body.find(qn("w:sectPr")).getprevious()
    if last is None or last.tag != qn("w:p"):
        return False
    if not "".join(last.itertext()).strip():
        return True
    return last.find(qn("w:pPr") + "/" + qn("w:outlineLvl")) is not None


def heading(doc, F, role, text, gap=True):
    """절·소절 제목. 앞에 빈 줄을 둔다(사용자 서론 점검본 2026-09-17) — 직전이 빈 줄이면 생략."""
    if gap and not _no_gap_needed(doc):
        emit(doc, F, "head_gap")
    return emit(doc, F, role, text)


def emit(doc, F, role, text=""):
    p = doc.add_paragraph()
    _set_ppr(p, F.pPr[role])
    if text:
        add_rich(p, text, F.rPr[role], sup=role in ("author", "affil"))
    return p


def figure(doc, F, name, num):
    """본문에서 그림을 인용한 자리에 이미지 + 캡션을 넣는다."""
    png = os.path.join(FIGDIR, name + ".png")
    if not os.path.exists(png):
        print(f"  ⚠️ 그림 없음: {name}.png")
        return False
    pic = doc.add_paragraph()
    _set_ppr(pic, F.pPr["fig_pic"])
    pic.add_run().add_picture(png, width=Inches(FIG_W.get(name, _FIG_DEFAULT_W)))
    cap = emit(doc, F, "fig_cap")
    add_run(cap, f"Fig. {num}. ", F.rPr["fig_cap"], bold=True)
    add_rich(cap, CAPTIONS[name], F.rPr["fig_cap"])
    return True


def md_table(doc, F, lines):
    rows = [[c.strip() for c in ln.strip().strip("|").split("|")] for ln in lines
            if not re.match(r"^\s*\|[\s:\-|]+\|\s*$", ln)]
    if not rows:
        return
    ncol = max(len(r) for r in rows)
    t = doc.add_table(rows=len(rows), cols=ncol)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, row in enumerate(rows):
        for j in range(ncol):
            p = t.cell(i, j).paragraphs[0]
            _set_ppr(p, F.cell_pPr)
            add_rich(p, row[j] if j < len(row) else "", F.cell_head if i == 0 else F.cell_body)
    _append_clone(doc, F.tbl_after)


def load_template():
    tpl = Document(TEMPLATE)
    F = Fmt(tpl)
    body = tpl.element.body
    for child in list(body):
        if child.tag != qn("w:sectPr"):
            body.remove(child)
    # 템플릿 본문에 딸려 있던 그림 관계는 이제 참조가 없다 — 남기면 파일만 커진다
    for rId, rel in list(tpl.part.rels.items()):
        if rel.reltype.endswith("/image"):
            tpl.part.drop_rel(rId)
    return tpl, F


def main():
    md = open(SRC, encoding="utf-8").read()
    doc, F = load_template()

    lines = md.split("\n")
    i, n_tbl, n_head, n_fig = 0, 0, 0, 0
    # 작업용 머리말(투고본에 실리지 않는 메타)과 ⚠️ 작업 메모는 제외한다(원본 md 에는 남는다)
    SKIP_META = re.compile(r"^\*\*(Target journal|Registration|Draft)\*\*")
    WORKNOTE = re.compile(r"⚠️|^To be completed before submission")
    CAPTION = re.compile(r"^\*\*(Table|Fig)")
    mode = None          # front · abstract · body · decl · refs · supp
    decl_started = False
    listish = 0

    while i < len(lines):
        st_ln = lines[i].strip()

        if st_ln.startswith("<!--"):                                 # md 내부 주석
            while i < len(lines) and "-->" not in lines[i]:
                i += 1
            i += 1
            continue
        if st_ln.startswith("# "):                                   # 논문 제목
            emit(doc, F, "title", st_ln[2:])
            i += 1
            continue
        if st_ln.startswith("## "):
            name = st_ln[3:].strip()
            if name == "Author information":        # 경계 마커 — 지면에 싣지 않는다
                mode = "front"
            elif name == "Abstract":
                _append_clone(doc, F.brk["abstract"])
                emit(doc, F, "abs_head", "ABSTRACT")
                mode = "abstract"; n_head += 1
            elif name == "Declarations":            # 경계 마커 — 결론 뒤 선언부
                _append_clone(doc, F.brk["decl"])
                mode, decl_started = "decl", False
            elif name == "References":
                emit(doc, F, "h1", name)
                mode = "refs"; n_head += 1
            elif name == "Supplementary material":
                _append_clone(doc, F.brk["supp"])
                emit(doc, F, "h1", name)
                mode = "supp"; n_head += 1
            else:
                first = mode == "abstract"
                if first:                           # 서론은 새 쪽에서 시작한다
                    _append_clone(doc, F.brk["intro"])
                heading(doc, F, "h1", name, gap=not first)
                mode = "body"; n_head += 1
            i += 1
            continue
        if st_ln.startswith("### "):
            n_head += 1
            heading(doc, F, "h2", st_ln[4:])
            i += 1
            continue
        if not st_ln or st_ln in ("---", "***") or SKIP_META.match(st_ln):
            i += 1
            continue
        if st_ln.startswith(">") or WORKNOTE.search(st_ln):          # 작업 메모·인용 블록
            i += 1
            continue
        if st_ln.startswith("|"):                                    # 표
            buf = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                buf.append(lines[i]); i += 1
            md_table(doc, F, buf)
            n_tbl += 1
            continue
        m = re.match(r"^<<FIG:(\w+)>>$", st_ln)                       # 본문 내 그림 삽입
        if m:
            n_fig += 1
            if not figure(doc, F, m.group(1), n_fig):
                n_fig -= 1
            i += 1
            continue

        buf = [st_ln]                                                 # 문단(줄바꿈 병합)
        i += 1
        if mode != "front":            # 전면부는 한 줄이 한 문단(저자·소속·주소)
            while i < len(lines) and lines[i].strip() and not re.match(
                    r"^\s*(#{1,6}\s|\||>|---|<<FIG:|<!--)", lines[i]):
                buf.append(lines[i].strip()); i += 1
        txt = " ".join(buf)

        if mode == "front":
            if SUPER.match(txt):                         # ¹ 소속
                emit(doc, F, "affil", txt)
                emit(doc, F, "fm_gap")
            elif txt.startswith("**["):                  # [Corresponding author]
                emit(doc, F, "corr_label", txt)
            elif re.match(r"^\\?\*Send", txt):           # 교신저자
                emit(doc, F, "corr", txt)
            elif SUPER.search(txt):                      # 저자¹
                emit(doc, F, "author", txt)
            else:                                        # 교신 주소 줄
                emit(doc, F, "address", txt)
        elif mode == "abstract":
            emit(doc, F, "keywords" if txt.startswith("**Keywords") else "abs_body", txt)
        elif mode == "decl":
            mh = re.fullmatch(r"\*\*(.+?)\*\*", txt)
            if mh:
                if decl_started:
                    emit(doc, F, "decl_gap")
                emit(doc, F, "h1", mh.group(1))
                decl_started = True
            else:
                emit(doc, F, "decl_body", txt)
        elif mode == "refs":
            emit(doc, F, "ref", txt)
        elif re.match(r"^\*\*RQ\d\.\*\*", txt):
            emit(doc, F, "rq", txt)
        elif CAPTION.match(txt):
            emit(doc, F, "tbl_cap", txt)
        else:
            if re.match(r"^([-*]|\d+\.)\s+", txt):
                listish += 1
            emit(doc, F, "body", txt)

    if listish:
        print(f"  ⚠️ 목록 형식 문단 {listish}개 — 서식본에 목록 견본이 없어 본문 문단으로 넣었다")

    # 본문에서 인용되지 않은 그림이 남았는지 — 조용히 빠지지 않게 한다
    used = set(re.findall(r"<<FIG:(\w+)>>", md))
    left = [k for k in CAPTIONS if k not in used]
    if left:
        print(f"  ⚠️ 본문에 삽입되지 않은 그림: {left}")

    # ★ 네이밍 규약(2026-08-17 사용자 확정): `Manuscript_{KO|EN}_{날짜}_ver{N}.docx`.
    #   같은 날짜의 기존 빌드(old/ 포함)와 충돌하지 않는 최소 N. HHMMSS 타임코드는
    #   "어느 게 최신인지 헷갈린다"는 사용자 지적으로 폐기.
    import glob as _glob
    tag = "KO" if "_KO" in os.path.basename(SRC) else "EN"
    day = datetime.datetime.now().strftime("%Y%m%d")
    outdir = os.path.join(ROOT, "01_논문작업")
    taken = set()
    for p in (_glob.glob(os.path.join(outdir, f"Manuscript_{tag}_{day}_ver*.docx"))
              + _glob.glob(os.path.join(outdir, "old", f"Manuscript_{tag}_{day}_ver*.docx"))):
        m = re.search(r"_ver(\d+)\.docx$", os.path.basename(p))
        if m:
            taken.add(int(m.group(1)))
    ver = 0
    while ver in taken:
        ver += 1
    out = os.path.join(outdir, f"Manuscript_{tag}_{day}_ver{ver}.docx")
    doc.save(out)
    # ★ 구판 자동 격리: 방금 만든 것을 뺀 같은 태그의 ver 빌드는 old/ 로 옮긴다
    #   (`_검토` 등 사용자 파일은 패턴에 안 걸려 안전). 현역 최종본은 항상 1개.
    os.makedirs(os.path.join(outdir, "old"), exist_ok=True)
    import shutil as _shutil
    for p in _glob.glob(os.path.join(outdir, f"Manuscript_{tag}_*_ver*.docx")):
        if os.path.abspath(p) != os.path.abspath(out):
            try:
                _shutil.move(p, os.path.join(outdir, "old", os.path.basename(p)))
                print(f"  [old/] ← {os.path.basename(p)}")
            except PermissionError:
                print(f"  ⚠️ {os.path.basename(p)} 이 Word 등에서 열려 있어 old/ 이관을 건너뛰었다 — 닫은 뒤 옮길 것")
    print(f"[완료] 문단 {len(doc.paragraphs)} · 표 {n_tbl} · 절제목 {n_head} · 그림 {n_fig}")
    print(f"[저장] {out}")
    return out


if __name__ == "__main__":
    main()
