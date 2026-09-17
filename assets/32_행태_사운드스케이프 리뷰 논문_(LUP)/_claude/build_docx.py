# -*- coding: utf-8 -*-
"""
Paper32 — 원고 markdown → Word(.docx)

★ 서식 정본 = `templates/Manuscript_KO_format_20260917_ver2.docx` (2026-09-17 사용자 ver6 편집본)
  사용자: "앞으로 이 양식을 베이스로 수정을 진행해야해. 내가 다시 편집할일 없게." 이 파일에서
  문단 서식뿐 아니라 ① 3선 표 서식(표마다 머리글·그룹 제목·그룹 첫/중간/끝 행·표 끝 행을 열별로 복제)
  ② 가로 쪽 구역(표마다 여백 포함) ③ 그림 삽입 폭 ④ 쪽 나눔 견본을 읽는다. md 쪽 표시:
  `<<PAGEBREAK>>` · `<<LANDSCAPE>>` … `<<END LANDSCAPE>>`(문서 끝까지 가로면 END 생략).
  그림 1·8 은 사용자가 `01_논문작업/Figure.pptx` 에서 다시 그린 것 → `figures/user/` 를 우선 쓴다.
(이전 정본 `templates/Manuscript_KO_format_20260917.docx` 의 설명)
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
from docx.shared import Inches, Emu
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(BASE)
SRC = os.path.join(ROOT, "01_논문작업",
                   sys.argv[1] if len(sys.argv) > 1 else "Manuscript_KO.md")
FIGDIR = os.path.join(BASE, "figures")
TEMPLATE = os.path.join(BASE, "templates", "Manuscript_KO_format_20260917_ver3.docx")
#   ver1(2026-09-17 서론 점검본): 본문 양쪽 정렬 · 절/소절 제목 앞 빈 줄
#   ver2(2026-09-17 사용자 ver6 편집본): 3선 표 · 가로 쪽(Table 2·4·B1) · 그림 1·8 재작도 · 서론 앞 쪽 나눔 없음
#   ver3(2026-09-17 사용자 ver7 최종 수정본, "내가 최종 수정한 버전을 가지고 향후 작업"): 교신저자 블록 복원
#        (단독·교신 저자, 저자명 뒤 ¹* 위첨자) · 3.8절·4장 앞 쪽 나눔 제거 · 3.4절 앞 빈 줄 없음(<<NOGAP>>)
USER_FIGDIR = os.path.join(FIGDIR, "user")     # 사용자가 다시 그린 그림(원본 = 01_논문작업/Figure.pptx)
# 열 폭을 서식본 값 대신 내용으로 다시 나누는 표 — 사용자 위임(2026-09-17 "table 3과 4 는 글이 길어서
# 폭을 조절하는데 한계"): 행 이름·MMAT 열을 줄이면서 폭도 새로 나눈다. 나머지 표는 사용자 폭 그대로.
REFLOW_TABLES = set()
#   2026-09-17 ver3: Table 3·4 도 사용자 ver7 의 열 폭(사용자가 Table 4 MMAT·근거 강도 열을 넓힘)을 그대로 쓴다.
#   표 내용이 크게 바뀌어 폭을 다시 나눠야 할 때만 {"Table 3", "Table 4"} 처럼 넣는다.

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


# ★ 2026-09-17 LUP 리뷰 예시(Zhang et al. 2025) 기준으로 재작성: 캡션은 그림이 무엇을 보여 주는지와
#   읽는 법(기호·색·단위·집계 규칙)만 쓴다. 결과 해석이나 강조("the clearest gap", "provides the
#   empirical basis")는 본문 결과·논의 절의 몫이라 캡션에서 뺐다.
CAPTIONS = {
    "Fig1_PRISMA": (
        "PRISMA 2020 flow diagram of study identification, screening and inclusion. Other methods "
        f"comprised backward and forward citation searching of the {_FD['prisma']['ct']['seeds']} "
        "studies included or retained for sensitivity analysis through database searching, and a "
        "supplementary search of OpenAlex for records not indexed in the three databases. Records from "
        "citation searching were prefiltered by an automated title-level filter before screening."),
    "Fig2_Forest": (
        "Forest plots of the four behavioural clusters: (a) walking speed, (b) staying or dwell time, "
        "(c) social interaction and (d) sound–behaviour correlation. Squares show the effect size of "
        "each study, and their size reflects the study's weight in the random-effects model (REML with "
        "the Hartung–Knapp adjustment); horizontal lines are 95% confidence intervals. Diamonds show the "
        "pooled estimate and its 95% confidence interval, and the dashed vertical line marks no effect. "
        "The columns on the right give each effect size with its 95% confidence interval and its "
        "weight (%). "
        "Effect sizes are Hedges' g in panels (a)–(c) and r in panel (d). Triangles mark studies "
        "retrieved by citation searching. Heterogeneity (I² and Q) and the p value of the pooled effect "
        "are given below each panel; prediction intervals and sensitivity analyses are given in Table 3."),
    "Fig3_EvidenceMap": (
        "Evidence map of behavioural domain by sound source. Each cell gives the number of included "
        "studies that examined the combination, and darker shading indicates more studies. A study "
        "contributes to every combination it examined, so cell counts do not sum to the number of "
        "included studies."),
    "Fig4_Direction": (
        "Direction of the relationship examined, by behavioural domain. Bars show the number of "
        "study × behavioural-domain records classified as forward (acoustic environment to behaviour), "
        "both, or reverse (behaviour or activity to acoustic environment or soundscape). The percentage "
        "to the right of each bar is the share of reverse records among forward and reverse records."),
    "Fig5_Methods": (
        "Behavioural measurement methods used in the included studies, by publication period. "
        "G1 = self-report; G2 = systematic observation; G3 = sensing, GPS, video or big-data "
        "measurement of behaviour. Studies that used more than one method are counted in each "
        "corresponding series."),
    "Fig6_Framework": (
        "Reciprocal evidence framework linking context, acoustic environment, soundscape "
        "appraisal, and observable behaviour. Spatial, physical, and socio-cultural context "
        "shapes the acoustic environment and moderates how acoustic conditions are interpreted "
        "and acted upon. Acoustic conditions may influence observable behaviour directly or "
        "through soundscape appraisal. Behaviour and activity can, in turn, modify the acoustic "
        "environment through occupancy and human sound production. Behavioural outcomes are "
        "organised below the framework along an engagement gradient from avoidance and passing to "
        "staying, interacting, and appropriating; this gradient is used as a synthesis device and "
        "is not a validated behavioural scale."),
    "Fig7_GeoTime": (
        f"Geographic and temporal distribution of the {_FD['n_included']} included studies. "
        "(a) Number of studies by country of data collection. Multi-country studies are counted once "
        "for each country, less frequent countries are grouped as Other (number of countries in "
        f"parentheses), and the {dict(_FD['geo']['countries']).get('Not reported', 0)} studies that did "
        "not report the country are not shown. (b) Number of studies by publication year and direction "
        "of the relationship examined; studies published before 2010 are combined in the first bar."),
    "Fig8_Quality": (
        "Methodological quality of the included studies appraised with MMAT 2018. (a) Number of "
        "studies rated high (4–5 criteria met), moderate (3) or low (0–2) in each MMAT study category. "
        "(b) Share of studies rated Yes, Can't tell or No on selected items, grouped by the proportion "
        "of studies meeting the criterion. Item numbers follow MMAT 2018: the first digit is the study "
        "category (1 = qualitative, 2 = randomised controlled trial, 3 = non-randomised, 4 = quantitative "
        "descriptive) and the second is the criterion within that category. "
        "The fraction to the right of each bar is the number of "
        "studies meeting the criterion over the number of studies to which the item applies. "
        "Can't tell indicates that the information needed to judge the criterion was not reported."),
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


def _plain(s):
    return re.sub(r"[*`]", "", s).strip()


def _text(el):
    return "".join(x.text or "" for x in el.iter(qn("w:t")))


def _landscape(sectPr):
    sz = sectPr.find(qn("w:pgSz")) if sectPr is not None else None
    return sz is not None and sz.get(qn("w:orient")) == "landscape"


def _text_width_in(sectPr):
    sz, mar = sectPr.find(qn("w:pgSz")), sectPr.find(qn("w:pgMar"))
    return (int(sz.get(qn("w:w"))) - int(mar.get(qn("w:left"))) - int(mar.get(qn("w:right")))) / 1440


class Fmt:
    """템플릿을 역할별 견본으로 읽어 둔다. 필수 견본을 못 찾으면 조용히 넘어가지 않고 멈춘다."""

    def __init__(self, tpl):
        P = tpl.paragraphs
        T = [p.text.strip() for p in P]

        def find(pred):
            return next((i for i, s in enumerate(T) if pred(s)), None)

        def at(pred, what):
            i = find(pred)
            if i is None:
                raise SystemExit(f"⚠️ 서식 템플릿에서 '{what}' 견본 문단을 찾지 못했다: {TEMPLATE}")
            return i

        i_abs = at(lambda s: s == "ABSTRACT", "ABSTRACT 제목")
        i_h1 = at(lambda s: s == "1. Introduction", "1수준 제목")
        i_h2 = at(lambda s: s.startswith("1.1 "), "2수준 제목")
        i_decl = at(lambda s: s == "CRediT authorship contribution statement", "선언부 제목")
        i_app = at(lambda s: s == "Supplementary material" or s.startswith("Appendix A"), "부록 제목")
        i_h1b = at(lambda s: s == "2. Methods", "2수준 절 제목")
        role = {
            "title": 0, "author": 1, "affil": 2, "fm_gap": 3,
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
        # 교신저자 블록은 단독 저자 확정(2026-09-17)으로 서식본에 없다 — 있으면 읽는다
        i_corr = find(lambda s: s.startswith("*Send correspondence"))
        if i_corr is not None:
            role.update(corr=i_corr, address=i_corr + 1)
        i_cl = find(lambda s: s == "[Corresponding author]")
        if i_cl is not None:
            role["corr_label"] = i_cl
        if role["body"] is None:
            raise SystemExit("⚠️ 서식 템플릿: 서론에서 양쪽 정렬된 본문 견본을 찾지 못했다")
        if T[role["head_gap"]]:
            raise SystemExit("⚠️ 서식 템플릿: 1.3 제목 앞 문단이 빈 줄이 아니다")
        self.pPr, self.rPr = {}, {}
        for k, i in role.items():
            pp = P[i]._p.pPr
            self.pPr[k] = _clean(pp) if pp is not None else None
            self.rPr[k] = _base_rpr(P[i])

        def is_pb(i):
            return i is not None and i >= 0 and any(
                b.get(qn("w:type")) == "page" for b in P[i]._p.iter(qn("w:br")))

        # 쪽 나눔 문단은 자리마다 견본이 조금씩 달라(사용자 편집 흔적) 자리별로 통째로 보관.
        # 서론 앞 쪽 나눔은 사용자가 ver6 에서 없앴다 — 서식본에 있을 때만 넣는다.
        self.brk = {k: _clean(P[i]._p) for k, i in
                    {"abstract": i_abs - 1, "intro": i_h1 - 1, "decl": i_decl - 1, "supp": i_app - 1}.items()
                    if is_pb(i)}
        for k in ("abstract", "decl", "supp"):
            if k not in self.brk:
                raise SystemExit(f"⚠️ 서식 템플릿: '{k}' 자리 앞 문단이 쪽 나눔이 아니다")
        mid = next((i for i in range(i_h1b, i_decl) if is_pb(i)), None)     # <<PAGEBREAK>> 견본
        self.pagebreak = _clean(P[mid]._p) if mid is not None else self.brk["decl"]
        pic = next(p for p in P if p._p.findall(".//" + qn("w:drawing")))
        self.pPr["fig_pic"] = _clean(pic._p.pPr)

        # ── 본문 요소를 차례로 훑어 표·구역·그림 폭 견본을 읽는다 ─────────────
        body = tpl.element.body
        els = list(body)
        self.tables = []          # (표 라벨, 머리글 서명, 표 요소)
        self.sect_portrait = None  # 세로 구역을 끝내는 문단(견본)
        self.sect_land = {}        # 표 라벨 → 그 표의 가로 구역을 끝내는 문단
        self.fig_w = {}            # 그림 이름 → 삽입 폭(EMU)
        label = None
        for k, el in enumerate(els):
            if el.tag == qn("w:tbl"):
                sig = tuple(_plain(_text(tc)) for tc in el.find(qn("w:tr")).findall(qn("w:tc")))
                self.tables.append((label, sig, _clean(el)))
                continue
            if el.tag != qn("w:p"):
                continue
            m = re.match(r"^(Table [A-Z]?\d+)\.", _text(el).strip())
            if m:
                label = m.group(1)
            sp = el.find(qn("w:pPr") + "/" + qn("w:sectPr"))
            if sp is not None:
                if _landscape(sp):
                    self.sect_land[label] = _clean(el)
                elif self.sect_portrait is None:
                    self.sect_portrait = _clean(el)
            if el.findall(".//" + qn("w:drawing")):
                cap = next((_text(x).strip() for x in els[k + 1:k + 3] if _text(x).strip()), "")
                mm = re.match(r"^Fig\. \d+\.\s*(.*)", cap)
                ext = el.find(".//" + qn("wp:extent"))
                if mm and ext is not None:
                    for name, c in CAPTIONS.items():
                        if mm.group(1)[:60] == c[:60]:
                            self.fig_w[name] = int(ext.get("cx"))
        final = body.find(qn("w:sectPr"))
        self.final = (label, _clean(final))       # 문서 끝 구역(서식본은 부록 B 표의 가로 구역)
        first_tbl = next(el for el in els if el.tag == qn("w:tbl"))
        nxt = first_tbl.getnext()
        if nxt is not None and _text(nxt).strip().startswith("Note."):
            nxt = nxt.getnext()
        if nxt is None or _text(nxt).strip() or nxt.find(".//" + qn("w:sectPr")) is not None:
            raise SystemExit("⚠️ 서식 템플릿: 첫 표(주석) 뒤에 빈 문단 견본이 없다")
        self.tbl_after = _clean(nxt)
        if not self.tables:
            raise SystemExit("⚠️ 서식 템플릿에 표 견본이 없다")

    def landscape_end(self, label):
        el = self.sect_land.get(label)
        if el is None:                   # lxml 요소는 참/거짓으로 판정하지 않는다(자식 수로 판정됨)
            el = next(iter(self.sect_land.values()), None)
        if el is None:
            raise SystemExit("⚠️ 서식 템플릿에 가로 구역 견본이 없다")
        return el

    def final_sectpr(self, landscape, label):
        if landscape:
            if self.final[0] == label and _landscape(self.final[1]):
                return deepcopy(self.final[1])
            return deepcopy(self.landscape_end(label).find(qn("w:pPr") + "/" + qn("w:sectPr")))
        if self.sect_portrait is not None:
            return deepcopy(self.sect_portrait.find(qn("w:pPr") + "/" + qn("w:sectPr")))
        return deepcopy(self.final[1])

    def table_width_in(self, tbl, label):
        """서식본 표의 전체 폭(인치): tblW 가 pct 면 그 표가 놓인 구역의 본문 폭 기준, dxa 면 그 값."""
        tw = tbl.find(qn("w:tblPr") + "/" + qn("w:tblW"))
        typ, val = (tw.get(qn("w:type")), int(tw.get(qn("w:w")) or 0)) if tw is not None else ("auto", 0)
        if typ == "dxa" and val:
            return val / 1440
        sect = (self.landscape_end(label).find(qn("w:pPr") + "/" + qn("w:sectPr"))
                if label in self.sect_land else self.final_sectpr(False, label))
        if typ == "pct" and val:
            return _text_width_in(sect) * val / 5000
        return _text_width_in(sect)


def _append_clone(doc, el):
    doc.element.body.find(qn("w:sectPr")).addprevious(deepcopy(el))


def _set_ppr(p, pPr):
    old = p._p.pPr
    if old is not None:
        p._p.remove(old)
    if pPr is not None:
        p._p.insert(0, deepcopy(pPr))


INLINE = re.compile(r"(\*\*.+?\*\*|(?<!\*)\*[^*\n]+?\*(?!\*)|`[^`]+?`)", re.S)
SUPER = re.compile(r"([¹²³⁴⁵⁶⁷⁸⁹⁰]+\*?)")     # 저자명 뒤 교신저자 표시 * 도 위첨자(사용자 ver7)
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


_NOGAP = [False]      # md <<NOGAP>> — 다음 제목 앞 빈 줄을 넣지 않는다(사용자 ver7: 3.4절 앞)


def heading(doc, F, role, text, gap=True):
    """절·소절 제목. 앞에 빈 줄을 둔다(사용자 서론 점검본 2026-09-17) — 직전이 빈 줄이면 생략."""
    if _NOGAP[0]:
        gap, _NOGAP[0] = False, False
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
    png = os.path.join(USER_FIGDIR, name + ".png")          # 사용자가 다시 그린 그림이 우선
    if not os.path.exists(png):
        png = os.path.join(FIGDIR, name + ".png")
    if not os.path.exists(png):
        print(f"  ⚠️ 그림 없음: {name}.png")
        return False
    pic = doc.add_paragraph()
    _set_ppr(pic, F.pPr["fig_pic"])
    width = Emu(F.fig_w[name]) if name in F.fig_w else Inches(FIG_W.get(name, _FIG_DEFAULT_W))
    pic.add_run().add_picture(png, width=width)
    cap = emit(doc, F, "fig_cap")
    add_run(cap, f"Fig. {num}. ", F.rPr["fig_cap"], bold=True)
    add_rich(cap, CAPTIONS[name], F.rPr["fig_cap"])
    return True


# 표 셀 글꼴 = 서식본 11 pt Times New Roman. 단어 폭은 시스템 글꼴 파일로 실측한다
#   (글자 수 × 어림값은 실제보다 13% 넓게 잡혀 표가 본문 폭을 넘는다고 오판했다 — ver6 실측).
#   글꼴 파일이 없는 기기에서는 어림값으로 대신한다.
_CELL_PT, _PAD_IN, _CHAR_IN = 11, 0.16, 0.069     # 여백 = Table Grid 기본 좌우 0.08 in
try:
    from PIL import ImageFont as _IF
    _FONT = {False: _IF.truetype("times.ttf", _CELL_PT * 10), True: _IF.truetype("timesbd.ttf", _CELL_PT * 10)}
except Exception:
    _FONT = None


def _text_in(s, bold=False):
    if _FONT is None:
        return len(s) * _CHAR_IN
    return _FONT[bold].getlength(s) / 10 / 72


def _reflow_widths(rows, ncol, total_in):
    """열 폭(인치)을 내용으로 나눈다. 좁은 열(숫자·짧은 값)은 한 줄에 들어가는 폭을 그대로 주고,
    남는 폭을 긴 텍스트 열에 한 줄 폭 비례로 나눈다(머리글은 단어 단위로 접혀도 된다).
    그룹 제목 행(첫 칸만 채운 행)은 셀을 합치므로 계산에서 뺀다."""
    split = lambda s: [x for x in re.split(r"(?<=[-/–])|\s+", s) if x]   # 하이픈·빗금 뒤도 줄바꿈 자리
    nat, mn = [], []
    for j in range(ncol):
        raw = [r[j] for r in rows[1:] if r[j] and not _is_group(r)]
        bold = [c.startswith("**") and c.endswith("**") for c in raw]
        cells = [_plain(c) for c in raw]
        head_tok = max([_text_in(x, bold=True) for x in split(_plain(rows[0][j]))] + [0.0])
        full = max([_text_in(c, bold=b) for c, b in zip(cells, bold)] + [0.0])
        tok = max([_text_in(x, bold=b) for c, b in zip(cells, bold) for x in split(c)] + [0.0])
        nat.append(max(full, head_tok) + _PAD_IN)
        mn.append(max(tok, head_tok) + _PAD_IN)
    if sum(nat) <= total_in:
        return [w * total_in / sum(nat) for w in nat]
    fixed, rest, remaining = {}, sorted(range(ncol), key=lambda j: nat[j]), total_in
    for j in list(rest):
        if nat[j] <= remaining / len(rest):
            fixed[j] = nat[j]
            remaining -= nat[j]
            rest.remove(j)
        else:
            break
    out = [fixed[j] if j in fixed else max(mn[j], remaining * nat[j] / sum(nat[k] for k in rest))
           for j in range(ncol)]
    over = sum(out) - total_in
    if over > 0:                     # 최소 폭 보장으로 넘치면 여유 있는 텍스트 열에서 덜어낸다
        room = {j: out[j] - mn[j] for j in rest if out[j] > mn[j]}
        for j, r in room.items():
            out[j] -= over * r / sum(room.values())
    return out


def _is_group(row):
    return len(row) > 1 and row[0].startswith("**") and not any(c.strip() for c in row[1:])


_TBLPR_ORDER = ["tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize", "tblStyleColBandSize",
                "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders", "shd", "tblLayout", "tblCellMar", "tblLook"]
_TCPR_ORDER = ["cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap", "tcMar",
               "textDirection", "tcFitText", "vAlign", "hideMark"]
_BRD_ORDER = ["top", "start", "left", "bottom", "end", "right", "insideH", "insideV", "tl2br", "tr2bl"]


def _local(el):
    return el.tag.split("}")[1]


def _ordered_insert(parent, child, order):
    idx = order.index(_local(child))
    for ex in parent:
        if _local(ex) in order and order.index(_local(ex)) > idx:
            ex.addprevious(child)
            return child
    parent.append(child)
    return child


def _sub(parent, name, order):
    el = parent.find(qn("w:" + name))
    return el if el is not None else _ordered_insert(parent, OxmlElement("w:" + name), order)


def _tcs(tr):
    return tr.findall(qn("w:tc"))


def _tpl_group(tr):
    tcs = _tcs(tr)
    span = tcs[0].find(qn("w:tcPr") + "/" + qn("w:gridSpan")) if len(tcs) == 1 else None
    return span is not None and int(span.get(qn("w:val"))) > 1


def _exemplars(tbl):
    """서식본 표의 행 견본: 머리글 · 첫 그룹 제목 · 그 밖의 그룹 제목 · 본문 행(그룹 첫/끝, 표 끝 여부별)."""
    trs = tbl.findall(qn("w:tr"))
    n = len(trs)
    ex = {"header": trs[0], "group_first": None, "group": None, "body": {}, "heights": []}
    for i in range(1, n):
        tr = trs[i]
        if _tpl_group(tr):
            key = "group_first" if i == 1 else "group"
            ex[key] = ex[key] if ex[key] is not None else tr
            continue
        flags = (i == 1 or _tpl_group(trs[i - 1]), i == n - 1 or _tpl_group(trs[i + 1]), i == n - 1)
        ex["body"].setdefault(flags, tr)
        h = tr.find(qn("w:trPr") + "/" + qn("w:trHeight"))
        if h is not None and h.get(qn("w:val")):
            ex["heights"].append(int(h.get(qn("w:val"))))
    ex["group_first"] = ex["group_first"] if ex["group_first"] is not None else ex["group"]
    ex["group"] = ex["group"] if ex["group"] is not None else ex["group_first"]
    return ex


def _copy_side(dst_tc, src_tc, side):
    """칸 테두리 한 변(top/bottom)을 다른 견본 칸의 값으로 바꾼다(견본에 없으면 지운다)."""
    src = src_tc.find(f"{qn('w:tcPr')}/{qn('w:tcBorders')}/{qn('w:' + side)}")
    tcPr = dst_tc.find(qn("w:tcPr"))
    if tcPr is None:
        tcPr = OxmlElement("w:tcPr")
        dst_tc.insert(0, tcPr)
    brd = tcPr.find(qn("w:tcBorders"))
    if brd is not None:
        for old in brd.findall(qn("w:" + side)):
            brd.remove(old)
    if src is not None:
        brd = brd if brd is not None else _ordered_insert(tcPr, OxmlElement("w:tcBorders"), _TCPR_ORDER)
        _ordered_insert(brd, deepcopy(src), _BRD_ORDER)


def _body_row(ex, first, last, tlast):
    """본문 행 견본. 같은 위치의 견본이 없으면(예: 한 행짜리 그룹) 중간 행을 바탕으로 위 테두리는
    그룹 첫 행에서, 아래 테두리는 같은 끝 조건의 행에서 가져와 조립한다."""
    b = ex["body"]
    if (first, last, tlast) in b:
        return deepcopy(b[(first, last, tlast)])
    base = deepcopy(b.get((False, False, False)) or next(iter(b.values())))
    top = next((v for k, v in b.items() if k[0] == first), None)
    bot = (next((v for k, v in b.items() if k[1:] == (last, tlast)), None)
           or next((v for k, v in b.items() if k[1] == last), None))
    for j, tc in enumerate(_tcs(base)):
        for side, src in (("top", top), ("bottom", bot)):
            if src is not None:
                s = _tcs(src)
                _copy_side(tc, s[min(j, len(s) - 1)], side)
    return base


def _fill(tc, text):
    """견본 칸의 문단 서식·런 서식은 두고 글자만 바꾼다."""
    ps = tc.findall(qn("w:p"))
    p = ps[0]
    for extra in ps[1:]:
        tc.remove(extra)
    rpr = next((deepcopy(r.find(qn("w:rPr"))) for r in p.findall(qn("w:r"))
                if r.find(qn("w:t")) is not None and r.find(qn("w:rPr")) is not None), None)
    for ch in list(p):
        if ch.tag != qn("w:pPr"):
            p.remove(ch)
    if rpr is not None and "**" in text:      # 원고에 굵게 표시가 있으면 그 표시만 따른다
        for x in rpr.findall(qn("w:b")) + rpr.findall(qn("w:bCs")):   # (사용자 ver6: "Setting †" 의 † 는 보통 글씨)
            rpr.remove(x)
    add_rich(Paragraph(p, None), text, rpr)


def _resize(tr, ncol):
    tcs = _tcs(tr)
    if len(tcs) == ncol:
        return
    for tc in tcs:
        tr.remove(tc)
    for j in range(ncol):
        tr.append(deepcopy(tcs[min(j, len(tcs) - 1)]))


def _apply_widths(tbl, widths_in):
    tw = [int(round(w * 1440)) for w in widths_in]
    grid = tbl.find(qn("w:tblGrid"))
    for gc in list(grid):
        grid.remove(gc)
    for w in tw:
        gc = OxmlElement("w:gridCol")
        gc.set(qn("w:w"), str(w))
        grid.append(gc)
    tblPr = tbl.find(qn("w:tblPr"))
    tblW = _sub(tblPr, "tblW", _TBLPR_ORDER)
    tblW.set(qn("w:w"), str(sum(tw)))
    tblW.set(qn("w:type"), "dxa")
    _sub(tblPr, "tblLayout", _TBLPR_ORDER).set(qn("w:type"), "fixed")
    for tr in tbl.findall(qn("w:tr")):
        trPr = tr.find(qn("w:trPr"))
        if trPr is not None:
            for tag in ("gridBefore", "gridAfter", "wBefore", "wAfter"):
                for x in trPr.findall(qn("w:" + tag)):
                    trPr.remove(x)
        tcs = _tcs(tr)
        pairs = [(tcs[0], sum(tw), len(tw))] if len(tcs) == 1 and len(tw) > 1 else \
                [(tc, w, 1) for tc, w in zip(tcs, tw)]
        for tc, w, span in pairs:
            tcPr = tc.find(qn("w:tcPr"))
            if tcPr is None:
                tcPr = OxmlElement("w:tcPr")
                tc.insert(0, tcPr)
            tcW = _sub(tcPr, "tcW", _TCPR_ORDER)
            tcW.set(qn("w:w"), str(w))
            tcW.set(qn("w:type"), "dxa")
            gs = tcPr.find(qn("w:gridSpan"))
            if span > 1:
                _sub(tcPr, "gridSpan", _TCPR_ORDER).set(qn("w:val"), str(span))
            elif gs is not None:
                tcPr.remove(gs)


def md_table(doc, F, lines, label, after=True):
    """서식본에서 같은 표(머리글 서명 → 표 라벨 순으로 찾음)를 견본으로 삼아 행 종류·열별로 복제한다.
    열 폭은 서식본 값을 그대로 쓰고, REFLOW_TABLES 와 견본이 없는 새 표만 내용으로 다시 나눈다."""
    rows = [[c.strip() for c in ln.strip().strip("|").split("|")] for ln in lines
            if not re.match(r"^\s*\|[\s:\-|]+\|\s*$", ln)]
    if not rows:
        return
    ncol = max(len(r) for r in rows)
    rows = [r + [""] * (ncol - len(r)) for r in rows]
    sig = tuple(_plain(c) for c in rows[0])
    tpl = next((e for lb, sg, e in F.tables if sg == sig), None)
    if tpl is None:
        tpl = next((e for lb, sg, e in F.tables if lb == label and len(sg) == ncol), None)
    generic = tpl is None
    if generic:
        tpl = F.tables[0][2]
        print(f"  ⚠️ {label}: 서식본에 같은 표가 없어 첫 표의 서식과 계산한 열 폭으로 만든다")
    ex = _exemplars(tpl)
    tbl = deepcopy(tpl)
    for tr in tbl.findall(qn("w:tr")):
        tbl.remove(tr)
    body_h = min(ex["heights"]) if ex["heights"] else None
    n = len(rows)
    for i, row in enumerate(rows):
        if i and _is_group(row):
            tr = deepcopy(ex["group_first"] if i == 1 else ex["group"])
            _fill(_tcs(tr)[0], row[0])
        else:
            if i == 0:
                tr = deepcopy(ex["header"])
            else:
                tr = _body_row(ex, i == 1 or _is_group(rows[i - 1]), i == n - 1 or _is_group(rows[i + 1]),
                               i == n - 1)
                h = tr.find(qn("w:trPr") + "/" + qn("w:trHeight"))
                if h is not None and body_h:       # 견본 행이 두 줄 높이여도 본문 행은 한 줄 높이에서 시작
                    h.set(qn("w:val"), str(body_h))
            if generic:
                _resize(tr, ncol)
            if len(_tcs(tr)) != ncol:
                raise SystemExit(f"⚠️ {label}: 서식본 표의 열 수({len(_tcs(tr))})와 원고 표({ncol})가 다르다")
            for tc, text in zip(_tcs(tr), row):
                _fill(tc, text)
        tbl.append(tr)
    if generic or label in REFLOW_TABLES:
        total = (F.table_width_in(tpl, label) if not generic
                 else _text_width_in(F.final_sectpr(False, label)))
        _apply_widths(tbl, _reflow_widths(rows, ncol, total))
    doc.element.body.find(qn("w:sectPr")).addprevious(tbl)
    if after:          # 표 주석이 이어지면 표 뒤 간격은 주석 다음에 둔다
        _append_clone(doc, F.tbl_after)


def _last(doc):
    return doc.element.body.find(qn("w:sectPr")).getprevious()


def _is_spacer(el):
    return (el is not None and el.tag == qn("w:p") and not _text(el).strip()
            and el.find(".//" + qn("w:br")) is None and el.find(".//" + qn("w:sectPr")) is None
            and el.find(".//" + qn("w:drawing")) is None)


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
    cur_label, landscape = None, False
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
                if mode == "front" and "address" in F.pPr and not _is_spacer(_last(doc)):
                    emit(doc, F, "fm_gap")                  # 교신 주소 뒤 빈 줄(사용자 ver7)
                _append_clone(doc, F.brk["abstract"])
                emit(doc, F, "abs_head", "ABSTRACT")
                mode = "abstract"; n_head += 1
            elif name == "Declarations":            # 경계 마커 — 결론 뒤 선언부
                _append_clone(doc, F.brk["decl"])
                mode, decl_started = "decl", False
            elif name == "References":
                emit(doc, F, "h1", name)
                mode = "refs"; n_head += 1
            elif name == "Supplementary material" or name.startswith("Appendix"):
                last = _last(doc)
                if last is not None and last.find(qn("w:pPr") + "/" + qn("w:sectPr")) is not None:
                    last.addprevious(deepcopy(F.brk["supp"]))   # 사용자 ver6 순서: 쪽 나눔 → 구역 나눔 → 제목
                else:
                    _append_clone(doc, F.brk["supp"])
                emit(doc, F, "h1", name)
                mode = "supp"; n_head += 1
            else:
                first = mode == "abstract"
                if first and "intro" in F.brk:      # 사용자 ver6: 서론 앞 쪽 나눔 없음
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
        if st_ln == "<<NOGAP>>":
            _NOGAP[0] = True
            i += 1
            continue
        if st_ln == "<<PAGEBREAK>>":
            _append_clone(doc, F.pagebreak)
            i += 1
            continue
        if st_ln == "<<LANDSCAPE>>":                                   # 여기서 세로 구역을 끝낸다
            if F.sect_portrait is None:
                raise SystemExit("⚠️ 서식 템플릿에 세로 구역 나눔 견본이 없다")
            _append_clone(doc, F.sect_portrait)
            landscape = True
            i += 1
            continue
        if st_ln == "<<END LANDSCAPE>>":                               # 여기서 가로 구역을 끝낸다
            if _is_spacer(_last(doc)):                                 # 표 주석 뒤 빈 줄 대신 구역 나눔
                doc.element.body.remove(_last(doc))
            _append_clone(doc, F.landscape_end(cur_label))
            landscape = False
            i += 1
            continue
        if st_ln.startswith("|"):                                    # 표
            buf = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                buf.append(lines[i]); i += 1
            k = i
            while k < len(lines) and not lines[k].strip():
                k += 1
            note_next = k < len(lines) and lines[k].strip().startswith("*Note.*")
            md_table(doc, F, buf, cur_label, after=not note_next)
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
                    r"^\s*(#{1,6}\s|\||>|---|<<|<!--)", lines[i]):
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
            m_lab = re.match(r"^\*\*(Table [A-Z]?\d+)\.\*\*", txt)
            if m_lab:
                cur_label = m_lab.group(1)
            emit(doc, F, "tbl_cap", txt)
        elif txt.startswith("*Note.*"):             # 표 주석(LUP 예시의 표 구성) — 캡션 서식
            emit(doc, F, "tbl_cap", txt)
            _append_clone(doc, F.tbl_after)
        else:
            if re.match(r"^([-*]|\d+\.)\s+", txt):
                listish += 1
            emit(doc, F, "body", txt)

    # 문서 끝 구역: 마지막이 가로 쪽이면 그 표의 가로 구역, 아니면 세로 구역
    body = doc.element.body
    old = body.find(qn("w:sectPr"))
    old.addprevious(F.final_sectpr(landscape, cur_label))
    body.remove(old)

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
