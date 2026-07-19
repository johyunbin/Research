# -*- coding: utf-8 -*-
# v8 최종본(20260630_211604) 전체 추출 — 재조립 기준선.
# 문단(스타일·그림 앵커 포함), 표 전체, 임베드 이미지 실측을 md로 덤프.
import os, re
from docx import Document
from docx.document import Document as _Doc
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

import sys
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
SRC = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "01_논문작업", "Manuscript_v8_20260630_211604.docx")
OUT = sys.argv[2] if len(sys.argv) > 2 else os.path.join(ROOT, "_claude", "review", "manuscript_v8_final_extract.md")

doc = Document(SRC)

def iter_blocks(parent):
    for child in parent.element.body.iterchildren():
        if child.tag == qn('w:p'):
            yield Paragraph(child, parent)
        elif child.tag == qn('w:tbl'):
            yield Table(child, parent)

lines = []
n_para = n_tbl = n_img = 0
img_info = []
fig_refs = set()

for blk in iter_blocks(doc):
    if isinstance(blk, Paragraph):
        n_para += 1
        # 그림(inline shape) 앵커 감지
        drawings = blk._element.findall('.//' + qn('w:drawing'))
        if drawings:
            for d in drawings:
                blip = d.find('.//' + qn('a:blip'))
                rid = blip.get(qn('r:embed')) if blip is not None else None
                part = doc.part.related_parts.get(rid) if rid else None
                name = os.path.basename(part.partname) if part else '?'
                size = len(part.blob) if part else 0
                n_img += 1
                img_info.append((n_img, name, size))
                lines.append(f"[[IMAGE {n_img}: {name} {size/1024:.0f}KB]]")
        t = blk.text
        if t.strip():
            style = blk.style.name if blk.style is not None else '?'
            bold = any(r.bold for r in blk.runs if r.bold)
            tag = "## " if (bold and len(t) < 120) else ""
            lines.append(f"{tag}{t}")
            for m in re.findall(r'Fig\.? ?\d+', t):
                fig_refs.add(m)
        elif not drawings:
            pass  # 빈 문단 생략
    else:
        n_tbl += 1
        lines.append(f"\n[[TABLE {n_tbl}]]")
        for row in blk.rows:
            cells = [c.text.replace('\n', ' / ') for c in row.cells]
            lines.append("  | " + " || ".join(cells))
        lines.append("")

hdr = [
    f"# v8 최종본 추출 ({os.path.basename(SRC)})",
    f"- 문단 {n_para} · 표 {n_tbl} · 임베드 이미지 {n_img}",
    f"- 이미지: " + ", ".join(f"#{i}:{n}({s/1024:.0f}KB)" for i, n, s in img_info),
    f"- 잔존 § 참조: {sum(l.count('§') for l in lines)}",
    "",
]
os.makedirs(os.path.dirname(OUT), exist_ok=True)
with open(OUT, "w", encoding="utf-8") as f:
    f.write("\n".join(hdr + lines))
print(f"문단 {n_para} · 표 {n_tbl} · 이미지 {n_img}")
for i, n, s in img_info:
    print(f"  IMG{i}: {n} {s/1024:.0f}KB")
print("저장:", OUT)
