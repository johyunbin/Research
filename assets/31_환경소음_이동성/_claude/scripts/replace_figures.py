# -*- coding: utf-8 -*-
# 사용자 수정 docx의 임베드 이미지 10개를 새 PNG로 제자리 교체(텍스트 수정 보존).
# width는 사용자 설정 유지, height만 새 종횡비로 재계산(왜곡 방지). 새 타임코드로 저장.
import os, struct
from datetime import datetime, timezone, timedelta
from docx import Document
from docx.shared import Emu

ROOT = r"T:\00_BACKUP\01_한양대학교\00_연구실\00_프로젝트\★_논문화 프로젝트\31_환경소음_이동성"
FIGD = os.path.join(ROOT, "_claude", "figures")
SRC = os.path.join(ROOT, "01_논문작업", "Manuscript_v7_20260621_173137.docx")
ts = datetime.now(timezone(timedelta(hours=9))).strftime("%Y%m%d_%H%M%S")
OUT = os.path.join(ROOT, "01_논문작업", f"Manuscript_v7_{ts}.docx")

FIGS = ["fig_study_area.png", "fig_identification_concept.png", "fig_segmented_forest.png",
        "fig_daynight.png", "fig_phase_mobility_maps.png", "fig_landuse_trajectory.png",
        "fig_did_eventstudy.png", "fig_change_map.png", "fig_spatial_bylanduse.png",
        "fig_drift_validation.png"]

def png_size(blob):
    assert blob[:8] == b"\x89PNG\r\n\x1a\n"
    w, h = struct.unpack(">II", blob[16:24])
    return w, h

doc = Document(SRC)
shapes = doc.inline_shapes
assert len(shapes) == len(FIGS), f"이미지 수 불일치 {len(shapes)} != {len(FIGS)}"
for i, shape in enumerate(shapes):
    blob = open(os.path.join(FIGD, FIGS[i]), "rb").read()
    wpx, hpx = png_size(blob)
    rId = shape._inline.graphic.graphicData.pic.blipFill.blip.embed
    doc.part.related_parts[rId]._blob = blob          # 이미지 데이터 교체
    cur_w = shape.width                                # 사용자 width 유지
    shape.height = Emu(int(round(int(cur_w) * hpx / wpx)))   # 새 종횡비로 height
    print(f"  Fig {i+1}: {FIGS[i]:32s} {wpx}x{hpx}px -> w={shape.width.inches:.2f} h={shape.height.inches:.2f}in")

doc.save(OUT)
print(f"\n저장: {OUT}")
print(f"(원본 보존: {os.path.basename(SRC)})")
