# -*- coding: utf-8 -*-
"""
Paper32 — 원고 markdown → Word(.docx)
서식은 paper31 `31_환경소음_이동성/01_논문작업/Manuscript_EN_20260727_005544.docx` 실측을 따른다:
  A4(8.27×11.69in) · 여백 L1.18/R·T·B 1.00in · Times New Roman
  제목 16pt bold · 본문·절제목 12pt(절제목 bold) · 표·표주 11pt(헤더 bold) · 줄간격 1.0
  줄번호(lnNumType) 켬 — Elsevier 요구
출력: 01_논문작업/Manuscript_EN_<타임코드>.docx
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
                   sys.argv[1] if len(sys.argv) > 1 else "Manuscript_EN_20260806_r2.md")
FIGDIR = os.path.join(BASE, "figures")

FONT = "Times New Roman"
FONT_EA = "바탕"   # 한글 본문(paper31 Normal 스타일과 동일 계열)
SZ_TITLE, SZ_BODY, SZ_TABLE = 16, 12, 11

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
        "Hartung–Knapp adjustment. Squares are individual effects, sized by inverse variance; "
        "diamonds are pooled estimates; triangles mark studies retrieved by citation searching. "
        f"Only social interaction (p = {_pl('social')['p']:.3f}) and the perception–behaviour "
        f"correlation (p = {_pl('correlation')['p']:.3f}) exclude zero, and both do so for the "
        "mean effect only — the 95% prediction interval includes zero in all four clusters."),
    "Fig3_EvidenceMap": (
        "Evidence map of behavioural domain by sound source. Cells count studies, and a study "
        "contributes to every cell it covers. Every combination is populated except aircraft "
        f"noise, which appears in {_FD['n_aircraft_records']} domain-level records from "
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
    "Fig6_Framework": (
        "Conceptual framework. The loop is bidirectional by construction: the forward path runs "
        "acoustic environment → appraisal → behaviour, the reverse path runs activity → sound "
        "production → acoustic environment. Behaviour is unfolded as an engagement gradient and "
        "each band carries the quantitative evidence attached to it. The framework's diagnostic "
        "value is visible in the leftmost band — the avoidance and walking-speed evidence, the "
        "most frequently cited claim in this literature, comes entirely from studies that MMAT "
        "rates low."),
    "Fig7_GeoTime": (
        f"Geographic and temporal distribution of the {_FD['n_included']} included studies. "
        f"(a) {_FD['geo']['countries'][0][1]} studies "
        f"({_FD['geo']['countries'][0][1] / _FD['n_included'] * 100:.0f}%) were conducted in "
        f"{_FD['geo']['countries'][0][0]}. (b) {_FD['n_since_2020']} studies "
        f"({_FD['n_since_2020'] / _FD['n_included'] * 100:.0f}%) appeared in 2020 or later, and "
        "reverse-direction studies appear only from the mid-2010s, so the bidirectional evidence "
        "base is younger still than the corpus as a whole."),
    "Fig8_Quality": (
        "MMAT 2018 appraisal. Two findings drive the Discussion. First, the two randomised "
        "studies score among the lowest rather than the highest: their reports omit the "
        "randomisation procedure, baseline comparability and blinding entirely, so the design "
        "cannot be credited. Second, the largest deficits are deficits of reporting — sample "
        f"representativeness (item 4.2, clearly met in {_Q['4.2']['Y']} of {_Q['4.2']['n']}), "
        f"low non-response bias (4.4, {_Q['4.4']['Y']} of {_Q['4.4']['n']}) and control of "
        f"confounding (3.4, {_Q['3.4']['Y']} of {_Q['3.4']['n']}) fail mostly because the "
        "information is absent, not because the studies are known to be biased."),
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
         space_after=0, space_before=0):
    p = doc.add_paragraph()
    pf = p.paragraph_format
    pf.line_spacing = 1.0
    pf.space_after = Pt(space_after)
    pf.space_before = Pt(space_before)
    if align is not None:
        p.alignment = align
    if indent is not None:
        pf.left_indent = Inches(indent)
    if text:
        add_rich(p, text, size, base_bold=bold)
    return p


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
    para(doc, "", SZ_TABLE)


def main():
    md = open(SRC, encoding="utf-8").read()
    doc = Document()

    st = doc.styles["Normal"]
    st.font.name = FONT
    st.font.size = Pt(SZ_BODY)
    st.paragraph_format.line_spacing = 1.0
    st.paragraph_format.space_after = Pt(0)
    rf = st.element.rPr.rFonts
    for a in ("w:ascii", "w:hAnsi", "w:cs"):
        rf.set(qn(a), FONT)
    rf.set(qn("w:eastAsia"), FONT_EA)

    s = doc.sections[0]
    s.page_width, s.page_height = Inches(8.27), Inches(11.69)
    s.left_margin = Inches(1.18)
    s.right_margin = s.top_margin = s.bottom_margin = Inches(1.00)
    enable_line_numbers(s)

    lines = md.split("\n")
    i, n_tbl, n_head = 0, 0, 0
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
            para(doc, st_ln[2:], SZ_TITLE, bold=True, space_after=12,
                 align=WD_ALIGN_PARAGRAPH.CENTER)
            i += 1; continue
        if st_ln == "## Author information":
            frontmatter = True
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
            para(doc, st_ln[4:], SZ_BODY, bold=True, space_before=10, space_after=4)
            i += 1; continue
        if st_ln.startswith("## "):                                  # 2수준 절
            n_head += 1
            para(doc, st_ln[3:], SZ_BODY, bold=True, space_before=12, space_after=5)
            i += 1; continue
        if st_ln in ("---", "***") or SKIP_META.match(st_ln):
            i += 1; continue
        if st_ln.startswith("> "):                                   # 인용 블록 → 들여쓴 본문
            buf = []
            while i < len(lines) and lines[i].strip().startswith(">"):
                buf.append(lines[i].strip().lstrip(">").strip()); i += 1
            txt = " ".join(buf)
            if not WORKNOTE.search(txt):
                para(doc, txt, SZ_TABLE, indent=0.3, space_after=6)
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
                para(doc, "• " + item, SZ_BODY, indent=0.25, space_after=2)
            para(doc, "", SZ_BODY, space_after=4)
            continue
        if not st_ln:
            i += 1; continue

        buf = [st_ln]                                                 # 본문 문단(줄바꿈 병합)
        i += 1
        while i < len(lines) and lines[i].strip() and not re.match(
                r"^\s*(#{1,6}\s|\||>|[-*]\s|\d+\.\s|---)", lines[i]):
            buf.append(lines[i].strip()); i += 1
        para(doc, " ".join(buf), SZ_BODY, space_after=6)

    # ── 그림 ──────────────────────────────────────────────────────
    doc.add_page_break()
    para(doc, "Figures", SZ_BODY, bold=True, space_after=8)
    n_fig = 0
    for k, name in enumerate(sorted(CAPTIONS), 1):
        png = os.path.join(FIGDIR, name + ".png")
        if not os.path.exists(png):
            print(f"  ⚠️ 그림 없음: {name}.png"); continue
        pic = doc.add_paragraph()
        pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pic.paragraph_format.space_before = Pt(12)
        pic.paragraph_format.space_after = Pt(3)
        pic.add_run().add_picture(png, width=Inches(5.6))
        cap = para(doc, "", SZ_TABLE, space_after=10)
        set_run(cap.add_run(f"Fig. {k}. "), SZ_TABLE, bold=True)
        add_rich(cap, CAPTIONS[name], SZ_TABLE)
        n_fig += 1

    ts = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    tag = "KO" if "_KO" in os.path.basename(SRC) else "EN"
    out = os.path.join(ROOT, "01_논문작업", f"Manuscript_{tag}_{ts}.docx")
    doc.save(out)
    print(f"[완료] 문단 {len(doc.paragraphs)} · 표 {n_tbl} · 절제목 {n_head} · 그림 {n_fig}")
    print(f"[저장] {out}")
    return out


if __name__ == "__main__":
    main()
