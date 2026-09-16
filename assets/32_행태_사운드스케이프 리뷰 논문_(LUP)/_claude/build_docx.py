# -*- coding: utf-8 -*-
"""
Paper32 — 원고 markdown → Word(.docx)
서식은 paper31 `31_환경소음_이동성/01_논문작업/Manuscript_EN_20260727_005544.docx` 실측을 따른다:
  A4(8.27×11.69in) · 여백 L1.18/R·T·B 1.00in · Times New Roman
  제목 16pt bold · 본문·절제목 12pt(절제목 bold) · 표·표주 11pt(헤더 bold) · 줄간격 1.0
  줄번호(lnNumType) 켬 — Elsevier 요구
출력: 01_논문작업/Manuscript_{KO|EN}_{YYYYMMDD}_ver{N}.docx (구판은 자동으로 old/ 이관)
"""
import sys, os, re, datetime

import docx
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.dirname(BASE)
SRC = os.path.join(ROOT, "01_논문작업",
                   sys.argv[1] if len(sys.argv) > 1 else "Manuscript_KO.md")
FIGDIR = os.path.join(BASE, "figures")

FONT = "Times New Roman"
FONT_EA = "바탕"   # 한글 본문(paper31 Normal 스타일과 동일 계열)
SZ_TITLE, SZ_BODY, SZ_TABLE = 16, 12, 11
LS_BODY, LS_HEAD = None, None    # 본문·제목 모두 Normal(2.0) 상속
#   ↑ 사용자가 직접 고친 11개 절 제목의 실측값(2026-08-10): JUSTIFY + 줄간격 미지정
LS_NORMAL = 2.0                  # Normal 스타일 줄간격
FIRST_INDENT = 0.2               # 본문 첫 줄 들여쓰기(in)
# 그림별 삽입 폭(in) — 세로로 긴 도면은 좁게 넣어야 한 쪽에 들어간다
# ★ 삽입 폭 = 작화 폭(viz_theme.W_FULL = 6.05 in) — 축소 없이 1:1 로 넣는다.
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
        f"Only social interaction (p = {_pl('social')['p']:.3f}) and the perception–behaviour "
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


def set_run(r, size=SZ_BODY, bold=False, italic=False):
    r.font.name = FONT
    r.font.size = Pt(size)
    r.font.bold = bold
    r.font.italic = italic
    # 동아시아 폰트도 같이 지정하지 않으면 한글·기호에서 다른 글꼴이 튄다
    rPr = r._element.get_or_add_rPr()
    rf = rPr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts"); rPr.append(rf)
    for a in ("w:ascii", "w:hAnsi", "w:cs"):
        rf.set(qn(a), FONT)
    rf.set(qn("w:eastAsia"), FONT_EA)
    return r


INLINE = re.compile(r"(\*\*.+?\*\*|(?<!\*)\*[^*\n]+?\*(?!\*)|`[^`]+?`)", re.S)


def add_rich(p, text, size=SZ_BODY, base_bold=False):
    """**bold** · *italic* · `code` 를 run 단위로 반영"""
    text = text.replace("\\*", "\u0001")
    for tok in INLINE.split(text):
        if not tok:
            continue
        tok = tok.replace("\u0001", "*")
        if tok.startswith("**") and tok.endswith("**") and len(tok) > 4:
            set_run(p.add_run(tok[2:-2]), size, bold=True)
        elif tok.startswith("*") and tok.endswith("*") and len(tok) > 2:
            set_run(p.add_run(tok[1:-1]), size, bold=base_bold, italic=True)
        elif tok.startswith("`") and tok.endswith("`") and len(tok) > 2:
            set_run(p.add_run(tok[1:-1]), size, bold=base_bold)
        else:
            set_run(p.add_run(tok), size, bold=base_bold)
    return p


def para(doc, text="", size=SZ_BODY, bold=False, align=None, indent=None,
         space_after=0, space_before=0, line_spacing=LS_BODY, first_indent=None):
    """paper31 실측 기본값: 본문 줄간격 2.0(더블) · 첫 줄 들여쓰기 0.2in."""
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.line_spacing = line_spacing
    pf.space_after = Pt(space_after)
    pf.space_before = Pt(space_before)
    if align is not None:
        p.alignment = align
    if indent is not None:
        pf.left_indent = Inches(indent)
    if first_indent is not None:
        pf.first_line_indent = Inches(first_indent)
    if text:
        add_rich(p, text, size, base_bold=bold)
    return p


def h2(doc, text):
    """2수준 제목 — Normal · JUSTIFY · bold · 검정(사용자 수정본 실측)."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf = p.paragraph_format
    pf.line_spacing = None; pf.space_before = None; pf.space_after = None
    pf.keep_with_next = True
    add_rich(p, text, SZ_BODY, base_bold=True)
    for r in p.runs:
        r.font.bold = True
        r.font.size = None                      # Normal(12pt) 상속 — 목표와 동일
        r.font.color.rgb = RGBColor(0, 0, 0)
    return p


def h3(doc, text):
    """3수준 제목 — Word `Heading 2` 스타일에 런을 TNR 12pt·bold 아님·검정으로 덮는다."""
    p = doc.add_paragraph(style="Heading 2")
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf = p.paragraph_format
    pf.line_spacing = None; pf.space_before = None; pf.space_after = None
    pf.keep_with_next = True                # 간격 10pt 는 Heading 2 스타일이 준다
    add_rich(p, text, SZ_BODY)
    for r in p.runs:
        r.font.bold = False
        r.font.color.rgb = RGBColor(0, 0, 0)
        r.font.size = Pt(SZ_BODY)
    return p


def figure(doc, name, num):
    """본문에서 그림을 인용한 자리에 이미지 + 캡션을 넣는다(paper31 배치와 동일)."""
    png = os.path.join(FIGDIR, name + ".png")
    if not os.path.exists(png):
        print(f"  ⚠️ 그림 없음: {name}.png"); return False
    pic = doc.add_paragraph()
    pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pic.paragraph_format.line_spacing = 1.0
    pic.paragraph_format.space_before = Pt(10)
    pic.paragraph_format.space_after = Pt(4)
    pic.add_run().add_picture(png, width=Inches(FIG_W.get(name, _FIG_DEFAULT_W)))
    cap = para(doc, "", SZ_TABLE, space_after=12, line_spacing=1.0)
    set_run(cap.add_run(f"Fig. {num}. "), SZ_TABLE, bold=True)
    add_rich(cap, CAPTIONS[name], SZ_TABLE)
    return True


def enable_line_numbers(section):
    sectPr = section._sectPr
    if sectPr.find(qn("w:lnNumType")) is None:
        ln = OxmlElement("w:lnNumType")
        ln.set(qn("w:countBy"), "1")
        ln.set(qn("w:restart"), "continuous")
        sectPr.append(ln)


def md_table(doc, lines):
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
            cell = t.cell(i, j)
            cell.text = ""
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.0
            p.paragraph_format.space_after = Pt(0)
            add_rich(p, row[j] if j < len(row) else "", SZ_TABLE, base_bold=(i == 0))
    para(doc, "", SZ_TABLE, line_spacing=1.0, space_after=6)


def main():
    md = open(SRC, encoding="utf-8").read()
    doc = Document()

    st = doc.styles["Normal"]
    st.font.name = FONT
    st.font.size = Pt(SZ_BODY)
    st.paragraph_format.line_spacing = LS_NORMAL
    st.paragraph_format.space_after = Pt(0)
    rf = st.element.rPr.rFonts
    for a in ("w:ascii", "w:hAnsi", "w:cs"):
        rf.set(qn(a), FONT)
    rf.set(qn("w:eastAsia"), FONT_EA)

    s = doc.sections[0]
    s.page_width, s.page_height = Inches(8.27), Inches(11.69)
    s.left_margin = Inches(1.18)
    s.right_margin = s.top_margin = s.bottom_margin = Inches(1.00)
    # 사용자가 줄번호를 껐다(2026-08-10 docx 실측) — 그 결정을 따른다.
    # enable_line_numbers(s)

    lines = md.split("\n")
    i, n_tbl, n_head, n_fig = 0, 0, 0, 0
    # 작업용 머리말(투고본에 실리지 않는 메타)은 제외한다
    SKIP_META = re.compile(r"^\*\*(Target journal|Registration|Draft)\*\*")
    # ⚠️ 로 시작하는 작업 메모는 투고본에 넣지 않는다(원본 md 에는 남는다)
    WORKNOTE = re.compile(r"⚠️|^To be completed before submission")
    # 전면부는 줄바꿈이 의미를 가진다(소속·교신저자 주소) — 문단 병합 금지
    frontmatter = False

    while i < len(lines):
        ln = lines[i]
        st_ln = ln.strip()

        if st_ln.startswith("# "):                                   # 논문 제목
            para(doc, st_ln[2:], SZ_TITLE, bold=True, space_before=14,
                 space_after=14, line_spacing=1.5)   # 좌측정렬(사용자 수정본 실측)
            i += 1; continue
        if st_ln == "## Author information":
            # 사용자 검토본(2026-08-16 확정 서식)은 이 제목을 지면에 싣지 않는다 —
            # md 에서는 전면부 블록의 경계 마커로만 쓴다.
            frontmatter = True
            i += 1; continue
        elif st_ln.startswith("## "):
            frontmatter = False
        if frontmatter and st_ln and not st_ln.startswith("#"):
            # 전면부는 줄바꿈이 의미를 갖되(소속·주소) 산문은 이어져야 한다.
            # `**라벨**`·위첨자·`\*` 로 시작하는 줄에서만 새 문단을 연다.
            if st_ln.startswith(">") or WORKNOTE.search(st_ln):
                i += 1; continue
            buf = [st_ln]; i += 1
            while i < len(lines):
                nxt = lines[i].strip()
                if (not nxt or nxt.startswith("#") or nxt.startswith(">")
                        or re.match(r"^(\*\*|[¹²³⁴]|\\\*)", nxt)):
                    break
                buf.append(nxt); i += 1
            para(doc, " ".join(buf).replace("\\*", "*"), SZ_BODY, space_after=4)
            continue
        if st_ln.startswith("### "):                                 # 3수준 절
            n_head += 1
            h3(doc, st_ln[4:])
            i += 1; continue
        if st_ln.startswith("## "):                                  # 2수준 절
            n_head += 1
            h2(doc, st_ln[3:])
            i += 1; continue
        if st_ln in ("---", "***") or SKIP_META.match(st_ln):
            i += 1; continue
        if st_ln.startswith("> "):                                   # 인용 블록 → 들여쓴 본문
            buf = []
            while i < len(lines) and lines[i].strip().startswith(">"):
                buf.append(lines[i].strip().lstrip(">").strip()); i += 1
            txt = " ".join(buf)
            if not WORKNOTE.search(txt):
                para(doc, txt, SZ_TABLE, indent=0.3, space_after=6, line_spacing=1.15)
            continue
        if st_ln.startswith("|"):                                    # 표
            buf = []
            while i < len(lines) and lines[i].strip().startswith("|"):
                buf.append(lines[i]); i += 1
            md_table(doc, buf); n_tbl += 1
            continue
        if re.match(r"^[-*]\s+", st_ln) or re.match(r"^\d+\.\s+", st_ln):   # 목록
            while i < len(lines) and (re.match(r"^[-*]\s+", lines[i].strip())
                                      or re.match(r"^\d+\.\s+", lines[i].strip())):
                item = re.sub(r"^([-*]|\d+\.)\s+", "", lines[i].strip())
                i += 1
                while i < len(lines) and lines[i].startswith("  ") and lines[i].strip() \
                        and not re.match(r"^\s*([-*]|\d+\.)\s+", lines[i]):
                    item += " " + lines[i].strip(); i += 1
                para(doc, item, SZ_BODY, indent=0.42, first_indent=-0.31,
                     space_before=0, space_after=8, line_spacing=LS_HEAD)
            continue
        if not st_ln:
            i += 1; continue

        m = re.match(r"^<<FIG:(\w+)>>$", st_ln)                       # 본문 내 그림 삽입
        if m:
            n_fig += 1
            if not figure(doc, m.group(1), n_fig):
                n_fig -= 1
            i += 1; continue

        buf = [st_ln]                                                 # 본문 문단(줄바꿈 병합)
        i += 1
        while i < len(lines) and lines[i].strip() and not re.match(
                r"^\s*(#{1,6}\s|\||>|[-*]\s|\d+\.\s|---|<<FIG:)", lines[i]):
            buf.append(lines[i].strip()); i += 1
        # 표/그림 캡션 문단은 들여쓰기 없이 단일 간격
        txt = " ".join(buf)
        if re.match(r"^\*\*(Table|Fig)", txt):
            para(doc, txt, SZ_TABLE, space_before=8, space_after=4, line_spacing=1.0)
        else:
            para(doc, txt, SZ_BODY, space_after=0, first_indent=FIRST_INDENT)

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
            _shutil.move(p, os.path.join(outdir, "old", os.path.basename(p)))
            print(f"  [old/] ← {os.path.basename(p)}")
    print(f"[완료] 문단 {len(doc.paragraphs)} · 표 {n_tbl} · 절제목 {n_head} · 그림 {n_fig}")
    print(f"[저장] {out}")
    return out


if __name__ == "__main__":
    main()
