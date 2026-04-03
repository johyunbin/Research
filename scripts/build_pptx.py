#!/usr/bin/env python3
"""
build_pptx.py — Build all 18 slides of the Samsung Research portfolio PPT.
Rewritten for RICH content: every slide fills the 13.333" x 7.5" space.
"""

import os
import sys

sys.path.insert(0, os.path.dirname(__file__))

from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx_helpers import (
    new_presentation, add_blank_slide, set_slide_bg,
    add_rect, add_accent_bar, add_textbox, add_multiline_textbox,
    add_section_title, add_metric_card, add_implication_box, add_image_safe,
    NAVY, ACCENT_BLUE, SKY_BLUE, DARK_TEXT, GRAY, GRAY2,
    LIGHT_BG, WARM_GRAY, WHITE, BLACK,
    SLIDE_WIDTH, SLIDE_HEIGHT, MARGIN_L, MARGIN_R, MARGIN_T,
    FONT_TITLE, FONT_BODY, FONT_EN,
)

ASSETS = os.path.join(os.path.dirname(os.path.dirname(__file__)), "assets")


def img(name: str) -> str:
    return os.path.join(ASSETS, name)


# ============================================================================
# S1: Title Slide
# ============================================================================
def build_s1(prs):
    slide = add_blank_slide(prs)
    set_slide_bg(slide, NAVY)

    # Accent blue bar near top
    add_accent_bar(slide, Inches(0.6), Inches(0.5), Inches(2.0),
                   Inches(0.06), color=ACCENT_BLUE)

    # Title
    add_textbox(slide, Inches(0.8), Inches(1.8), Inches(11.5), Inches(1.8),
                "Spatial Audio Research &\nPerception-driven Quality Evaluation",
                font_size=32, font_color=WHITE, bold=True,
                alignment=PP_ALIGN.CENTER, font_name=FONT_EN)

    # Name
    add_textbox(slide, Inches(0.8), Inches(3.9), Inches(11.5), Inches(0.5),
                "조현인 (Hyun In Jo, Ph.D.)",
                font_size=14, font_color=WHITE, bold=True,
                alignment=PP_ALIGN.CENTER)

    # Subtitle
    add_textbox(slide, Inches(0.8), Inches(4.5), Inches(11.5), Inches(0.4),
                "Senior Research Engineer, Hyundai Motor Company (NVH Division)",
                font_size=11, font_color=WARM_GRAY, bold=False,
                alignment=PP_ALIGN.CENTER)

    # Org
    add_textbox(slide, Inches(0.8), Inches(5.0), Inches(11.5), Inches(0.4),
                "Samsung Research · Visual Technology · Display Innovation Lab · Spatial Audio",
                font_size=11, font_color=WARM_GRAY, bold=False,
                alignment=PP_ALIGN.CENTER)

    # Contact
    add_textbox(slide, Inches(0.8), Inches(5.5), Inches(11.5), Inches(0.4),
                "best2012@naver.com | 010-6387-8402 | linkedin.com/in/hyunin-jo",
                font_size=10, font_color=WARM_GRAY, bold=False,
                alignment=PP_ALIGN.CENTER, font_name=FONT_EN)


# ============================================================================
# S2: About Me
# ============================================================================
def build_s2(prs):
    slide = add_blank_slide(prs)

    add_section_title(slide, MARGIN_L, MARGIN_T, Inches(12),
                      "About Me",
                      "조현인 (Hyun In Jo, Ph.D.) — 경력 · 핵심 실적 · 전문성")

    # --- Left: Career Timeline ---
    tl_left = Inches(0.8)
    tl_top = Inches(1.8)
    entries = [
        ("2013-2016", "B.S. 건축공학, 한양대 (수석졸업, 조기졸업)"),
        ("2016-2022", "Ph.D. 건축음향, 한양대 (석박통합, GPA 4.39/4.5)"),
        ("2022.03-08", "Post-doc, 한국건설기술연구원"),
        ("2022.08-현재", "현대자동차 NVH 책임연구원"),
        ("NOW →", "Samsung Research, Spatial Audio"),
    ]
    for i, (year, desc) in enumerate(entries):
        y = tl_top + Inches(i * 0.7)
        add_rect(slide, tl_left, y + Inches(0.08), Inches(0.12), Inches(0.12),
                 fill_color=NAVY)
        if i < len(entries) - 1:
            add_rect(slide, tl_left + Inches(0.05), y + Inches(0.2),
                     Inches(0.02), Inches(0.5), fill_color=ACCENT_BLUE)
        add_textbox(slide, tl_left + Inches(0.3), y, Inches(1.5), Inches(0.3),
                    year, font_size=10, font_color=ACCENT_BLUE, bold=True,
                    font_name=FONT_EN)
        add_textbox(slide, tl_left + Inches(1.9), y, Inches(3.5), Inches(0.4),
                    desc, font_size=10, font_color=DARK_TEXT)

    # --- Right: 4 Metric Cards ---
    card_left = Inches(7.0)
    card_w = Inches(2.7)
    card_h = Inches(0.85)
    cards = [
        ("SCI(E) 24편", "주저자 21편, h-index 18"),
        ("EAA Best Paper", "ICA 2019, I-INCE Young Professional"),
        ("특허 6건", "국내+미국, 기술이전 5천만원"),
        ("국제공동연구", "UCL·소르본, SATP 18개국 표준화"),
    ]
    for i, (num, lbl) in enumerate(cards):
        cy = Inches(1.8) + Inches(i * 1.05)
        add_metric_card(slide, card_left, cy, card_w, card_h,
                        num, lbl, number_size=18, label_size=9)

    # --- Bottom: 4 Competency Boxes ---
    box_w = Inches(2.85)
    box_h = Inches(0.7)
    box_top = Inches(6.2)
    parts = [
        ("Part I", "Spatial Audio &\nImmersive Rendering"),
        ("Part II", "Perception-driven\nQuality Evaluation"),
        ("Part III", "Research-to-Product\nExecution"),
        ("Part IV", "AI-driven Audio\nProcessing"),
    ]
    for i, (label, desc) in enumerate(parts):
        bx = Inches(0.6) + Inches(i * 3.1)
        add_rect(slide, bx, box_top, box_w, box_h, fill_color=NAVY)
        add_textbox(slide, bx + Inches(0.1), box_top + Inches(0.05),
                    box_w - Inches(0.2), Inches(0.2),
                    label, font_size=9, font_color=ACCENT_BLUE, bold=True,
                    font_name=FONT_EN)
        add_textbox(slide, bx + Inches(0.1), box_top + Inches(0.25),
                    box_w - Inches(0.2), Inches(0.4),
                    desc, font_size=10, font_color=WHITE, bold=False)


# ============================================================================
# S3: Key Question
# ============================================================================
def build_s3(prs):
    slide = add_blank_slide(prs)

    add_textbox(slide, Inches(0.8), Inches(0.5), Inches(11.5), Inches(1.0),
                "THD 0.01%, 주파수 응답 ±0.5dB —\n공학 스펙이 완벽해도 사용자가 \"좋다\"고 느끼지 않을 수 있다",
                font_size=16, font_color=GRAY, bold=False,
                alignment=PP_ALIGN.CENTER)

    add_textbox(slide, Inches(0.8), Inches(1.6), Inches(11.5), Inches(0.5),
                "시각 맥락만으로 오디오 만족도가 76% 좌우된다면?",
                font_size=14, font_color=ACCENT_BLUE, bold=True,
                alignment=PP_ALIGN.CENTER)

    ctr_w = Inches(9.0)
    ctr_h = Inches(2.2)
    ctr_l = Inches((13.333 - 9.0) / 2)
    ctr_t = Inches(2.4)
    add_rect(slide, ctr_l, ctr_t, ctr_w, ctr_h, fill_color=NAVY)
    add_textbox(slide, ctr_l, ctr_t, ctr_w, ctr_h,
                "사용자가 진짜 몰입을 느끼는\n3D Audio-Visual 경험을\n어떻게 설계하고 검증할 것인가?",
                font_size=22, font_color=WHITE, bold=True,
                alignment=PP_ALIGN.CENTER,
                anchor=MSO_ANCHOR.MIDDLE)

    card_w = Inches(2.6)
    card_h = Inches(1.2)
    card_top = Inches(5.2)
    parts = [
        "Part I\nSpatial Audio &\nImmersive Rendering",
        "Part II\nPerception-driven\nQuality Evaluation",
        "Part III\nResearch-to-Product\nExecution",
        "Part IV\nAI-driven Audio\nProcessing",
    ]
    for i, txt in enumerate(parts):
        cx = Inches(0.5) + Inches(i * 3.2)
        add_rect(slide, cx, card_top, card_w, card_h, fill_color=LIGHT_BG)
        add_textbox(slide, cx + Inches(0.1), card_top + Inches(0.1),
                    card_w - Inches(0.2), card_h - Inches(0.2),
                    txt, font_size=10, font_color=NAVY, bold=True,
                    alignment=PP_ALIGN.CENTER,
                    anchor=MSO_ANCHOR.MIDDLE)
        if i < 3:
            ax = cx + card_w
            add_textbox(slide, ax, card_top + Inches(0.4), Inches(0.6), Inches(0.4),
                        "→", font_size=18, font_color=ACCENT_BLUE, bold=True,
                        alignment=PP_ALIGN.CENTER, font_name=FONT_EN)


# ============================================================================
# S4: Soundscape Introduction
# ============================================================================
def build_s4(prs):
    slide = add_blank_slide(prs)

    add_section_title(slide, MARGIN_L, MARGIN_T, Inches(12),
                      "사운드스케이프란?",
                      "Noise Control → Soundscape 패러다임 전환")

    lw = Inches(5.0)
    lh = Inches(1.5)
    lt = Inches(1.8)
    add_rect(slide, Inches(0.6), lt, lw, lh, fill_color=LIGHT_BG)
    add_textbox(slide, Inches(0.8), lt + Inches(0.15), lw - Inches(0.4), Inches(0.3),
                "Traditional Noise Control", font_size=13, font_color=NAVY, bold=True)
    add_textbox(slide, Inches(0.8), lt + Inches(0.55), lw - Inches(0.4), Inches(0.7),
                '"소음이 얼마나 큰가?" dB 측정',
                font_size=11, font_color=DARK_TEXT)

    add_textbox(slide, Inches(5.8), lt + Inches(0.4), Inches(1.2), Inches(0.5),
                "→", font_size=28, font_color=ACCENT_BLUE, bold=True,
                alignment=PP_ALIGN.CENTER, font_name=FONT_EN)

    add_rect(slide, Inches(7.2), lt, lw, lh, fill_color=NAVY)
    add_textbox(slide, Inches(7.4), lt + Inches(0.15), lw - Inches(0.4), Inches(0.3),
                "New Paradigm: Soundscape", font_size=13, font_color=WHITE, bold=True)
    add_textbox(slide, Inches(7.4), lt + Inches(0.55), lw - Inches(0.4), Inches(0.7),
                '"소리가 어떻게 경험되는가?" 인간 지각 중심',
                font_size=11, font_color=WHITE)

    add_rect(slide, Inches(0.6), Inches(3.6), Inches(11.8), Inches(0.8), fill_color=LIGHT_BG)
    add_textbox(slide, Inches(0.8), Inches(3.7), Inches(11.4), Inches(0.6),
                "ISO 12913: \"acoustic environment as perceived or experienced and/or understood "
                "by a person or people, in context\"",
                font_size=11, font_color=GRAY, bold=False,
                alignment=PP_ALIGN.CENTER)

    add_image_safe(slide, img("proto_s5_img2.png"),
                   Inches(3.5), Inches(4.5), Inches(6.0), Inches(2.7),
                   placeholder_text="Pleasant-Eventful Diagram")


# ============================================================================
# S5: Soundscape → Spatial Audio Bridge
# ============================================================================
def build_s5(prs):
    slide = add_blank_slide(prs)

    add_section_title(slide, MARGIN_L, MARGIN_T, Inches(12),
                      "왜 사운드스케이프가 Spatial Audio에 직결되는가",
                      "")

    rows = [
        ("소리를 경험으로 평가\n(dB가 아닌 사용자 지각 중심)",
         "렌더링 품질을 THD·주파수응답이 아닌\n사용자가 느끼는 공간감·몰입감으로 평가"),
        ("오디오-비주얼 상호작용\n(시각 맥락이 청각 지각을 최대 76% 좌우)",
         "Display + Audio 통합 설계\nHolographic Displays × Spatial Audio 시너지"),
        ("재생 환경에 따라 동일 음원 지각 변화",
         "거실·침실·차량 등 재생 공간별\n렌더링 최적화"),
        ("개인차 (소음 민감도·성격·청력)",
         "Customized Audio 개인화"),
        ("대규모 지각 평가 프로토콜\n(ISO 12913 + SATP 18개국, 134명)",
         "Eclipsa Audio 품질 벤치마킹 및\n인증 기준 수립"),
    ]

    col_w = Inches(5.2)
    row_h = Inches(0.85)
    left1 = Inches(0.6)
    left2 = Inches(7.0)
    arrow_l = Inches(5.9)

    for i, (left_t, right_t) in enumerate(rows):
        ry = Inches(1.5) + Inches(i * 0.95)
        add_rect(slide, left1, ry, col_w, row_h, fill_color=LIGHT_BG)
        add_textbox(slide, left1 + Inches(0.15), ry + Inches(0.08),
                    col_w - Inches(0.3), row_h - Inches(0.16),
                    left_t, font_size=10, font_color=DARK_TEXT)
        add_textbox(slide, arrow_l, ry + Inches(0.15), Inches(1.0), Inches(0.5),
                    "→", font_size=16, font_color=ACCENT_BLUE, bold=True,
                    alignment=PP_ALIGN.CENTER, font_name=FONT_EN)
        add_rect(slide, left2, ry, col_w, row_h, fill_color=LIGHT_BG)
        add_textbox(slide, left2 + Inches(0.15), ry + Inches(0.08),
                    col_w - Inches(0.3), row_h - Inches(0.16),
                    right_t, font_size=10, font_color=DARK_TEXT)

    bar_top = Inches(6.5)
    add_rect(slide, Inches(0.6), bar_top, Inches(12.1), Inches(0.7), fill_color=NAVY)
    add_textbox(slide, Inches(0.8), bar_top + Inches(0.1), Inches(11.7), Inches(0.5),
                '신호처리가 "어떻게 구현할 것인가"라면, 저의 전문성은 '
                '"사용자가 어떻게 경험할 것인가"를 설계하고 검증하는 것입니다',
                font_size=11, font_color=WHITE, bold=True,
                alignment=PP_ALIGN.CENTER)


# ============================================================================
# S6: Spatial Audio Overview — VISUAL RESEARCH MAP
# ============================================================================
def build_s6(prs):
    slide = add_blank_slide(prs)

    # Title area
    add_section_title(slide, MARGIN_L, MARGIN_T, Inches(12),
                      "Part I. Spatial Audio 기술 역량 — 연구 체계도",
                      "5개 기술 축 + Eclipsa Audio 연결 맵")

    # --- Central hub: large navy pill ---
    hub_w = Inches(4.0)
    hub_h = Inches(0.8)
    hub_l = Inches((13.333 - 4.0) / 2)
    hub_t = Inches(1.65)
    add_rect(slide, hub_l, hub_t, hub_w, hub_h, fill_color=NAVY)
    add_textbox(slide, hub_l, hub_t, hub_w, hub_h,
                "Spatial Audio Perception Research",
                font_size=15, font_color=WHITE, bold=True,
                alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                font_name=FONT_EN)

    # --- 5 axis cards: 2 rows (3 top, 2 bottom) for better space usage ---
    axes = [
        ("축1", "시각 재현 방식", "APAC'22 (IF 3.4)", "HMD vs Monitor\n40명 × 8환경", "p < 0.05", ACCENT_BLUE),
        ("축2", "재생 방식", "APAC'19 (IF 3.4)", "헤드폰 vs 스피커\n4 test환경 × 6 levels", "2.2 dBA 차이", ACCENT_BLUE),
        ("축3", "바이노럴-비주얼", "B&E'19 ×2 (IF 7.4)", "HRTF × HMD 2×2\n40명 × 320 data", "HRTF 77%", SKY_BLUE),
        ("축4", "시청각 정보", "B&E'20 (IF 7.4)", "Audio/Visual/AV\n3조건 × 8환경", "Visual 76%", SKY_BLUE),
        ("축5", "생태학적 타당성", "SCS'21 (IF 11.7)", "VR vs In-situ\n50명 × 10환경", "유의차 없음", RGBColor(0x0A, 0x7A, 0xB5)),
    ]

    # Row 1: axes 1-3
    row1_top = Inches(2.8)
    card_w = Inches(3.7)
    card_h = Inches(1.85)

    for idx, (ax_label, ax_title, ax_ref, ax_detail, ax_metric, color) in enumerate(axes[:3]):
        cx = Inches(0.5) + Inches(idx * 4.1)
        # Vertical connecting line from hub
        line_x = cx + card_w / 2 - Inches(0.01)
        add_rect(slide, line_x, hub_t + hub_h, Inches(0.025), row1_top - hub_t - hub_h,
                 fill_color=color)
        # Card background
        add_rect(slide, cx, row1_top, card_w, card_h, fill_color=LIGHT_BG,
                 line_color=color, line_width_pt=1.5)
        # Colored top strip
        add_rect(slide, cx, row1_top, card_w, Inches(0.06), fill_color=color)
        # Axis label
        add_textbox(slide, cx + Inches(0.15), row1_top + Inches(0.12), Inches(0.8), Inches(0.25),
                    ax_label, font_size=11, font_color=color, bold=True, font_name=FONT_EN)
        # Title
        add_textbox(slide, cx + Inches(0.9), row1_top + Inches(0.12), card_w - Inches(1.1), Inches(0.25),
                    ax_title, font_size=12, font_color=NAVY, bold=True)
        # Reference
        add_textbox(slide, cx + Inches(0.15), row1_top + Inches(0.4), card_w - Inches(0.3), Inches(0.2),
                    ax_ref, font_size=9, font_color=GRAY, font_name=FONT_EN)
        # Detail
        add_textbox(slide, cx + Inches(0.15), row1_top + Inches(0.65), card_w - Inches(0.3), Inches(0.6),
                    ax_detail, font_size=10, font_color=DARK_TEXT)
        # Metric badge
        add_rect(slide, cx + Inches(0.15), row1_top + Inches(1.35), card_w - Inches(0.3), Inches(0.38),
                 fill_color=color)
        add_textbox(slide, cx + Inches(0.15), row1_top + Inches(1.35), card_w - Inches(0.3), Inches(0.38),
                    ax_metric, font_size=13, font_color=WHITE, bold=True,
                    alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font_name=FONT_EN)

    # Row 2: axes 4-5, centered
    row2_top = Inches(5.0)
    for idx, (ax_label, ax_title, ax_ref, ax_detail, ax_metric, color) in enumerate(axes[3:]):
        cx = Inches(1.5) + Inches(idx * 5.5)
        # Connecting line
        line_x = cx + card_w / 2 - Inches(0.01)
        # Diagonal feel — just use vertical from hub bottom
        add_rect(slide, line_x, hub_t + hub_h, Inches(0.025),
                 row2_top - hub_t - hub_h, fill_color=color)
        # Card
        add_rect(slide, cx, row2_top, card_w, card_h, fill_color=LIGHT_BG,
                 line_color=color, line_width_pt=1.5)
        add_rect(slide, cx, row2_top, card_w, Inches(0.06), fill_color=color)
        add_textbox(slide, cx + Inches(0.15), row2_top + Inches(0.12), Inches(0.8), Inches(0.25),
                    ax_label, font_size=11, font_color=color, bold=True, font_name=FONT_EN)
        add_textbox(slide, cx + Inches(0.9), row2_top + Inches(0.12), card_w - Inches(1.1), Inches(0.25),
                    ax_title, font_size=12, font_color=NAVY, bold=True)
        add_textbox(slide, cx + Inches(0.15), row2_top + Inches(0.4), card_w - Inches(0.3), Inches(0.2),
                    ax_ref, font_size=9, font_color=GRAY, font_name=FONT_EN)
        add_textbox(slide, cx + Inches(0.15), row2_top + Inches(0.65), card_w - Inches(0.3), Inches(0.6),
                    ax_detail, font_size=10, font_color=DARK_TEXT)
        add_rect(slide, cx + Inches(0.15), row2_top + Inches(1.35), card_w - Inches(0.3), Inches(0.38),
                 fill_color=color)
        add_textbox(slide, cx + Inches(0.15), row2_top + Inches(1.35), card_w - Inches(0.3), Inches(0.38),
                    ax_metric, font_size=13, font_color=WHITE, bold=True,
                    alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font_name=FONT_EN)

    # --- Eclipsa Audio connection box (right side, between rows) ---
    ecl_l = Inches(9.8)
    ecl_t = Inches(5.0)
    ecl_w = Inches(3.2)
    ecl_h = Inches(1.85)
    add_rect(slide, ecl_l, ecl_t, ecl_w, ecl_h, fill_color=NAVY)
    add_multiline_textbox(
        slide, ecl_l + Inches(0.15), ecl_t + Inches(0.15), ecl_w - Inches(0.3), ecl_h - Inches(0.3),
        [
            ("→ Eclipsa Audio", True, WHITE, 14),
            ("", False, WHITE, 6),
            ("HRTF 개인화 (77%)", False, WHITE, 10),
            ("A-V 통합 설계 (76%)", False, WHITE, 10),
            ("VR 품질 평가 실증", False, WHITE, 10),
            ("생리반응 품질 검증", False, WHITE, 10),
            ("92.5% 만족도 예측", False, WHITE, 10),
        ],
        font_color=WHITE,
        line_spacing=1.25,
    )
    # Arrow pointing to Eclipsa box
    add_textbox(slide, Inches(9.2), Inches(5.5), Inches(0.6), Inches(0.5),
                "→", font_size=22, font_color=ACCENT_BLUE, bold=True,
                alignment=PP_ALIGN.CENTER, font_name=FONT_EN)


# ============================================================================
# S7: Axis 1 - Visual Reproduction
# ============================================================================
def build_s7(prs):
    slide = add_blank_slide(prs)

    # Title
    add_section_title(slide, MARGIN_L, MARGIN_T, Inches(7),
                      "축1: 시각 재현 방식",
                      "")

    # Paper reference (top right)
    add_rect(slide, Inches(8.0), Inches(0.5), Inches(5.0), Inches(0.55), fill_color=LIGHT_BG)
    add_multiline_textbox(
        slide, Inches(8.15), Inches(0.55), Inches(4.7), Inches(0.4),
        [
            ("APAC_2022_Jo & Jeon", True, NAVY, 11),
            ("Applied Acoustics, IF 3.4", False, GRAY, 9),
        ],
        line_spacing=1.2,
    )

    # RQ box (full width)
    rq_top = Inches(1.25)
    add_rect(slide, Inches(0.6), rq_top, Inches(12.1), Inches(0.65), fill_color=LIGHT_BG)
    add_textbox(slide, Inches(0.8), rq_top + Inches(0.08), Inches(11.7), Inches(0.5),
                "RQ: 같은 Spatial Audio를 HMD vs 모니터로 볼 때, 사용자의 음환경 지각이 얼마나 달라지는가?",
                font_size=12, font_color=NAVY, bold=True)

    # === LEFT COLUMN: Methodology ===
    left_x = Inches(0.6)
    meth_top = Inches(2.1)
    meth_w = Inches(5.8)

    add_textbox(slide, left_x, meth_top, meth_w, Inches(0.3),
                "Methodology", font_size=13, font_color=NAVY, bold=True)

    add_multiline_textbox(
        slide, left_x, meth_top + Inches(0.35), meth_w, Inches(2.2),
        [
            ("실험설계: 40명 피험자 × 8개 도시 음환경", True, DARK_TEXT, 11),
            ("시각 재현: HMD (HTC VIVE Pro) vs 2D Monitor", False, DARK_TEXT, 10),
            ("오디오: First-Order Ambisonics (FOA) + 실시간 Head-tracking", False, DARK_TEXT, 10),
            ("평가: 14개 semantic differential pairs", False, DARK_TEXT, 10),
            ("통계: 반복측정 ANOVA, Bonferroni 사후검정", False, DARK_TEXT, 10),
            ("", False, DARK_TEXT, 6),
            ("핵심 발견:", True, ACCENT_BLUE, 11),
            ("  - HMD 조건: 공간 현실감(Presence) 유의하게 증가 (p < 0.05)", False, DARK_TEXT, 10),
            ("  - 모니터 조건: 전반적 소음 인식(Overall awareness) 더 높음", False, DARK_TEXT, 10),
            ("  - 시각 재현 방식이 오디오 품질 판단을 체계적으로 변화시킴", False, DARK_TEXT, 10),
        ],
        line_spacing=1.3,
    )

    # === RIGHT COLUMN: Figure ===
    fig_x = Inches(6.7)
    fig_top = Inches(2.1)
    add_textbox(slide, fig_x, fig_top, Inches(6.0), Inches(0.3),
                "Result — Semantic Profile Comparison", font_size=11, font_color=GRAY, bold=True)
    add_image_safe(slide, img("apac2022_p7.png"),
                   fig_x, fig_top + Inches(0.35), Inches(6.0), Inches(3.0),
                   placeholder_text="apac2022_p7.png")

    # === BOTTOM: Metric cards + Implication ===
    mc_top = Inches(5.6)
    mc_h = Inches(0.85)
    mc_w = Inches(2.5)
    metrics = [
        ("40 × 8", "피험자 × 환경"),
        ("p < 0.05", "HMD Presence 유의차"),
        ("14 pairs", "Semantic Differential"),
    ]
    for i, (num, lbl) in enumerate(metrics):
        mx = Inches(0.6) + Emu(i * (int(mc_w) + int(Inches(0.25))))
        add_metric_card(slide, mx, mc_top, mc_w, mc_h, num, lbl,
                        number_size=22, label_size=9)

    # Implication box
    add_implication_box(slide, Inches(8.2), mc_top, Inches(4.5), Inches(1.6),
                        [
                            ("Samsung 시사점", True, WHITE, 12),
                            ("XR 디바이스 대응 시, 시각 조건 통제가", False, WHITE, 10),
                            ("오디오 품질 평가의 필수 전제 조건", False, WHITE, 10),
                            ("→ Eclipsa Audio 품질 평가 프로토콜에", False, WHITE, 10),
                            ("  시각 재현 방식 변수 반드시 포함", False, WHITE, 10),
                        ],
                        font_size=10, line_spacing=1.25)


# ============================================================================
# S8: Axis 2 - Headphone vs Speaker
# ============================================================================
def build_s8(prs):
    slide = add_blank_slide(prs)

    add_section_title(slide, MARGIN_L, MARGIN_T, Inches(7),
                      "축2: 헤드폰 vs 스피커",
                      "")

    # Paper reference
    add_rect(slide, Inches(8.0), Inches(0.5), Inches(5.0), Inches(0.55), fill_color=LIGHT_BG)
    add_multiline_textbox(
        slide, Inches(8.15), Inches(0.55), Inches(4.7), Inches(0.4),
        [
            ("APAC_2019_Jeon et al", True, NAVY, 11),
            ("Applied Acoustics, IF 3.4", False, GRAY, 9),
        ],
        line_spacing=1.2,
    )

    # RQ
    rq_top = Inches(1.25)
    add_rect(slide, Inches(0.6), rq_top, Inches(12.1), Inches(0.65), fill_color=LIGHT_BG)
    add_textbox(slide, Inches(0.8), rq_top + Inches(0.08), Inches(11.7), Inches(0.5),
                "RQ: 헤드폰과 스피커 재생 방식이 동일 음원에 대한 사용자의 소리 품질 판단을 어떻게 바꾸는가?",
                font_size=12, font_color=NAVY, bold=True)

    # === LEFT: Methodology ===
    left_x = Inches(0.6)
    meth_top = Inches(2.1)
    meth_w = Inches(5.8)

    add_textbox(slide, left_x, meth_top, meth_w, Inches(0.3),
                "Methodology", font_size=13, font_color=NAVY, bold=True)

    add_multiline_textbox(
        slide, left_x, meth_top + Inches(0.35), meth_w, Inches(2.2),
        [
            ("4가지 재생 조건:", True, DARK_TEXT, 11),
            ("  (1) Headphone only  (2) Speaker only", False, DARK_TEXT, 10),
            ("  (3) Headphone + HMD  (4) Speaker + HMD", False, DARK_TEXT, 10),
            ("자극: LAeq 40-65 dB, 6단계 소음 레벨", False, DARK_TEXT, 10),
            ("종속변수: 성가심(Annoyance), 허용한계(Allowance)", False, DARK_TEXT, 10),
            ("", False, DARK_TEXT, 6),
            ("핵심 발견:", True, ACCENT_BLUE, 11),
            ("  - Annoyance 차이: headphone vs speaker 8%", False, DARK_TEXT, 10),
            ("  - Allowance 차이: 6%", False, DARK_TEXT, 10),
            ("  - 50% annoyance level에서 SPL 차이: 2.2 dBA", False, DARK_TEXT, 10),
            ("  - Speaker + HMD 조합 = 실제 현장에 가장 근접한 반응", False, DARK_TEXT, 10),
        ],
        line_spacing=1.25,
    )

    # === RIGHT: Figure ===
    fig_x = Inches(6.7)
    fig_top = Inches(2.1)
    add_textbox(slide, fig_x, fig_top, Inches(6.0), Inches(0.3),
                "Result — Annoyance vs SPL by Condition", font_size=11, font_color=GRAY, bold=True)
    add_image_safe(slide, img("be2019a_p7.png"),
                   fig_x, fig_top + Inches(0.35), Inches(6.0), Inches(3.0),
                   placeholder_text="be2019a_p7.png")

    # === BOTTOM: Metrics + Implication ===
    mc_top = Inches(5.6)
    mc_h = Inches(0.85)
    mc_w = Inches(2.0)
    metrics = [
        ("8%", "Annoyance 차이"),
        ("6%", "Allowance 차이"),
        ("2.2 dBA", "50% SPL 차이"),
    ]
    for i, (num, lbl) in enumerate(metrics):
        mx = Inches(0.6) + Emu(i * (int(mc_w) + int(Inches(0.25))))
        add_metric_card(slide, mx, mc_top, mc_w, mc_h, num, lbl,
                        number_size=24, label_size=9)

    # Implication
    add_implication_box(slide, Inches(7.5), mc_top, Inches(5.2), Inches(1.6),
                        [
                            ("Samsung 시사점", True, WHITE, 12),
                            ("헤드폰(바이노럴) vs 스피커(멀티채널) 재생 시", False, WHITE, 10),
                            ("사용자 지각 차이가 체계적으로 발생", False, WHITE, 10),
                            ("→ Eclipsa Audio 렌더링 품질 평가에서", False, WHITE, 10),
                            ("  재생 디바이스별 보정 기준 필요", False, WHITE, 10),
                            ("→ Speaker+HMD가 reference condition으로 적합", False, WHITE, 10),
                        ],
                        font_size=10, line_spacing=1.2)


# ============================================================================
# S9: Axis 3 - Binaural-Visual Interaction
# ============================================================================
def build_s9(prs):
    slide = add_blank_slide(prs)

    add_section_title(slide, MARGIN_L, MARGIN_T, Inches(7),
                      "축3: 바이노럴-비주얼 상호작용",
                      "")

    add_rect(slide, Inches(8.0), Inches(0.5), Inches(5.0), Inches(0.55), fill_color=LIGHT_BG)
    add_multiline_textbox(
        slide, Inches(8.15), Inches(0.55), Inches(4.7), Inches(0.4),
        [
            ("B&E_2019 x 2편", True, NAVY, 11),
            ("Building and Environment, IF 7.4", False, GRAY, 9),
        ],
        line_spacing=1.2,
    )

    rq_top = Inches(1.25)
    add_rect(slide, Inches(0.6), rq_top, Inches(12.1), Inches(0.65), fill_color=LIGHT_BG)
    add_textbox(slide, Inches(0.8), rq_top + Inches(0.08), Inches(11.7), Inches(0.5),
                "RQ: HRTF 바이노럴 렌더링과 HMD 시각 재현, 어느 것이 사용자 공간 지각에 더 지배적인가?",
                font_size=12, font_color=NAVY, bold=True)

    # === LEFT: Methodology ===
    left_x = Inches(0.6)
    meth_top = Inches(2.1)
    meth_w = Inches(5.8)

    add_textbox(slide, left_x, meth_top, meth_w, Inches(0.3),
                "Methodology", font_size=13, font_color=NAVY, bold=True)

    add_multiline_textbox(
        slide, left_x, meth_top + Inches(0.35), meth_w, Inches(2.3),
        [
            ("2x2 Factorial Design:", True, DARK_TEXT, 11),
            ("  Factor A: HRTF (Individualized vs Generic)", False, DARK_TEXT, 10),
            ("  Factor B: HMD (VR 360 vs No Visual)", False, DARK_TEXT, 10),
            ("피험자: 40명 × 8개 도시 음환경 = 320 data points", False, DARK_TEXT, 10),
            ("장비: Sennheiser HD-650 + HTC VIVE Pro", False, DARK_TEXT, 10),
            ("HRTF: CIPIC HRTF Database", False, DARK_TEXT, 10),
            ("", False, DARK_TEXT, 6),
            ("핵심 발견:", True, ACCENT_BLUE, 11),
            ("  - HRTF가 공간감 지각의 77% 설명 (지배적 요인)", False, DARK_TEXT, 10),
            ("  - HMD는 23% 보조적 역할", False, DARK_TEXT, 10),
            ("  - HRTF+HMD 동시 적용 시 음상 외재화·몰입감 유의 증가", False, DARK_TEXT, 10),
            ("  - VR 환경에서 허용한계 6~7 dB 하락 (더 민감한 반응)", False, DARK_TEXT, 10),
        ],
        line_spacing=1.2,
    )

    # === RIGHT: Figure ===
    fig_x = Inches(6.7)
    fig_top = Inches(2.1)
    add_textbox(slide, fig_x, fig_top, Inches(6.0), Inches(0.3),
                "Result — HRTF vs HMD Contribution", font_size=11, font_color=GRAY, bold=True)
    add_image_safe(slide, img("be2019a_p9.png"),
                   fig_x, fig_top + Inches(0.35), Inches(6.0), Inches(3.0),
                   placeholder_text="be2019a_p9.png")

    # === BOTTOM: Metrics + Implication ===
    mc_top = Inches(5.6)
    mc_h = Inches(0.85)
    mc_w = Inches(2.0)
    metrics = [
        ("77%", "HRTF 공간감 기여"),
        ("23%", "HMD 시각 기여"),
        ("6~7 dB", "VR 허용한계 하락"),
    ]
    for i, (num, lbl) in enumerate(metrics):
        mx = Inches(0.6) + Emu(i * (int(mc_w) + int(Inches(0.25))))
        add_metric_card(slide, mx, mc_top, mc_w, mc_h, num, lbl,
                        number_size=26, label_size=9)

    add_implication_box(slide, Inches(7.5), mc_top, Inches(5.2), Inches(1.6),
                        [
                            ("Samsung 시사점", True, WHITE, 12),
                            ("HRTF 개인화 = 렌더러 최적화의 최우선 과제 (77%)", False, WHITE, 10),
                            ("→ Eclipsa Audio HRTF 개인화 알고리즘 개발 시", False, WHITE, 10),
                            ("  공간감 향상 효과가 시각 재현보다 3배 이상 지배적", False, WHITE, 10),
                            ("→ 제한된 리소스에서 HRTF에 집중 투자 근거 제공", False, WHITE, 10),
                        ],
                        font_size=10, line_spacing=1.2)


# ============================================================================
# S10: Axis 4 - Audio-Visual Information
# ============================================================================
def build_s10(prs):
    slide = add_blank_slide(prs)

    add_section_title(slide, MARGIN_L, MARGIN_T, Inches(7),
                      "축4: 시청각 정보 기여도",
                      "")

    add_rect(slide, Inches(8.0), Inches(0.5), Inches(5.0), Inches(0.55), fill_color=LIGHT_BG)
    add_multiline_textbox(
        slide, Inches(8.15), Inches(0.55), Inches(4.7), Inches(0.4),
        [
            ("B&E_2020_Jeon & Jo", True, NAVY, 11),
            ("Building and Environment, IF 7.4", False, GRAY, 9),
        ],
        line_spacing=1.2,
    )

    rq_top = Inches(1.25)
    add_rect(slide, Inches(0.6), rq_top, Inches(12.1), Inches(0.65), fill_color=LIGHT_BG)
    add_textbox(slide, Inches(0.8), rq_top + Inches(0.08), Inches(11.7), Inches(0.5),
                "RQ: Audio 정보와 Visual 정보가 전체 환경 만족도에 각각 얼마나 기여하는가?",
                font_size=12, font_color=NAVY, bold=True)

    # === LEFT: Methodology ===
    left_x = Inches(0.6)
    meth_top = Inches(2.1)
    meth_w = Inches(5.8)

    add_textbox(slide, left_x, meth_top, meth_w, Inches(0.3),
                "Methodology", font_size=13, font_color=NAVY, bold=True)

    add_multiline_textbox(
        slide, left_x, meth_top + Inches(0.35), meth_w, Inches(2.3),
        [
            ("오디오: FOA Ambisonics + HMD + Head-tracking", True, DARK_TEXT, 11),
            ("환경: 8개 도시 음환경 (공원, 도로, 광장 등)", False, DARK_TEXT, 10),
            ("3가지 제시 조건:", True, DARK_TEXT, 11),
            ("  (1) Audio-only  (2) Visual-only  (3) Audio+Visual", False, DARK_TEXT, 10),
            ("종속변수: 만족도, 자연스러움, 쾌적성", False, DARK_TEXT, 10),
            ("", False, DARK_TEXT, 6),
            ("핵심 발견:", True, ACCENT_BLUE, 11),
            ("  - 만족도 기여: Audio 24% vs Visual 76%", False, DARK_TEXT, 10),
            ("  - 만족도 예측 모델 설명력: R² = 51%", False, DARK_TEXT, 10),
            ("  - Audio가 landscape 자연스러움에도 영향 (cross-modal effect)", False, DARK_TEXT, 10),
            ("  - 시각이 지배적이나, 오디오 없이는 자연스러움 저하", False, DARK_TEXT, 10),
        ],
        line_spacing=1.2,
    )

    # === RIGHT: Figure ===
    fig_x = Inches(6.7)
    fig_top = Inches(2.1)
    add_textbox(slide, fig_x, fig_top, Inches(6.0), Inches(0.3),
                "Result — A/V Contribution to Satisfaction", font_size=11, font_color=GRAY, bold=True)
    add_image_safe(slide, img("be2020_p9.png"),
                   fig_x, fig_top + Inches(0.35), Inches(6.0), Inches(3.0),
                   placeholder_text="be2020_p9.png")

    # === BOTTOM: Metrics + Implication ===
    mc_top = Inches(5.6)
    mc_h = Inches(0.85)
    mc_w = Inches(2.0)
    metrics = [
        ("24%", "Audio 만족도 기여"),
        ("76%", "Visual 만족도 기여"),
        ("R²=51%", "만족도 모델 설명력"),
    ]
    for i, (num, lbl) in enumerate(metrics):
        mx = Inches(0.6) + Emu(i * (int(mc_w) + int(Inches(0.25))))
        add_metric_card(slide, mx, mc_top, mc_w, mc_h, num, lbl,
                        number_size=26, label_size=9)

    add_implication_box(slide, Inches(7.5), mc_top, Inches(5.2), Inches(1.6),
                        [
                            ("Samsung 시사점", True, WHITE, 12),
                            ("Display + Audio 통합 설계가 만족도의 핵심", False, WHITE, 10),
                            ("→ Holographic Displays x Spatial Audio 시너지:", False, WHITE, 10),
                            ("  시각 76% + 오디오 24%의 cross-modal 효과 극대화", False, WHITE, 10),
                            ("→ 오디오 품질만 올려도 자연스러움 개선 가능", False, WHITE, 10),
                        ],
                        font_size=10, line_spacing=1.2)


# ============================================================================
# S11: Axis 5 - Ecological Validity
# ============================================================================
def build_s11(prs):
    slide = add_blank_slide(prs)

    add_section_title(slide, MARGIN_L, MARGIN_T, Inches(7),
                      "축5: 생태학적 타당성",
                      "")

    add_rect(slide, Inches(8.0), Inches(0.5), Inches(5.0), Inches(0.55), fill_color=LIGHT_BG)
    add_multiline_textbox(
        slide, Inches(8.15), Inches(0.55), Inches(4.7), Inches(0.4),
        [
            ("SCS_2021_Jo & Jeon", True, NAVY, 11),
            ("Sustainable Cities and Society, IF 11.7", False, GRAY, 9),
        ],
        line_spacing=1.2,
    )

    rq_top = Inches(1.25)
    add_rect(slide, Inches(0.6), rq_top, Inches(12.1), Inches(0.65), fill_color=LIGHT_BG)
    add_textbox(slide, Inches(0.8), rq_top + Inches(0.08), Inches(11.7), Inches(0.5),
                "RQ: VR 실험실에서의 음환경 평가를 실제 현장(in-situ) 결과와 동일하게 신뢰할 수 있는가?",
                font_size=12, font_color=NAVY, bold=True)

    # === LEFT: Methodology ===
    left_x = Inches(0.6)
    meth_top = Inches(2.1)
    meth_w = Inches(5.8)

    add_textbox(slide, left_x, meth_top, meth_w, Inches(0.3),
                "Methodology", font_size=13, font_color=NAVY, bold=True)

    add_multiline_textbox(
        slide, left_x, meth_top + Inches(0.35), meth_w, Inches(2.3),
        [
            ("대규모 실험: 50명 피험자 × 10개 도시 음환경", True, DARK_TEXT, 11),
            ("ISO 12913-2 표준 3개 프로토콜 비교:", True, DARK_TEXT, 11),
            ("  Method A: 현장 직접 평가 (In-situ)", False, DARK_TEXT, 10),
            ("  Method B: 현장 녹음 후 실험실 재생", False, DARK_TEXT, 10),
            ("  Method C: VR (FOA Ambisonics + HMD + Head-tracking)", False, DARK_TEXT, 10),
            ("", False, DARK_TEXT, 6),
            ("핵심 발견:", True, ACCENT_BLUE, 11),
            ("  - VR vs In-situ: 통계적 유의차 없음 (생태학적 타당성 실증)", False, DARK_TEXT, 10),
            ("  - Pleasantness-Eventful 모델이 3개 프로토콜 모두에서 재현", False, DARK_TEXT, 10),
            ("  - Method C에서 비음향적 요인(시각, 맥락) 발견", False, DARK_TEXT, 10),
            ("  - 가장 높은 IF(11.7) — 연구 영향력 입증", False, DARK_TEXT, 10),
        ],
        line_spacing=1.2,
    )

    # === RIGHT: Figure ===
    fig_x = Inches(6.7)
    fig_top = Inches(2.1)
    add_textbox(slide, fig_x, fig_top, Inches(6.0), Inches(0.3),
                "Result — Protocol Comparison (P-E Model)", font_size=11, font_color=GRAY, bold=True)
    add_image_safe(slide, img("scs2021_p4.png"),
                   fig_x, fig_top + Inches(0.35), Inches(6.0), Inches(3.0),
                   placeholder_text="scs2021_p4.png")

    # === BOTTOM: Metrics + Implication ===
    mc_top = Inches(5.6)
    mc_h = Inches(0.85)
    mc_w = Inches(2.0)
    metrics = [
        ("50 × 10", "피험자 × 환경"),
        ("3 Protocols", "ISO 12913-2 A/B/C"),
        ("VR ≈ In-situ", "유의차 없음"),
    ]
    for i, (num, lbl) in enumerate(metrics):
        mx = Inches(0.6) + Emu(i * (int(mc_w) + int(Inches(0.25))))
        add_metric_card(slide, mx, mc_top, mc_w, mc_h, num, lbl,
                        number_size=22, label_size=9)

    add_implication_box(slide, Inches(7.5), mc_top, Inches(5.2), Inches(1.6),
                        [
                            ("Samsung 시사점", True, WHITE, 12),
                            ("VR 실험실에서 Eclipsa Audio 렌더링 품질 평가 →", False, WHITE, 10),
                            ("실사용 환경 결과와 동일하게 신뢰 가능", False, WHITE, 10),
                            ("→ 제품 출시 전 VR 기반 대규모 평가 프레임워크", False, WHITE, 10),
                            ("  구축 가능 (현장 테스트 비용 대폭 절감)", False, WHITE, 10),
                        ],
                        font_size=10, line_spacing=1.2)


# ============================================================================
# S12: Multimodal Biosignal Modeling
# ============================================================================
def build_s12(prs):
    slide = add_blank_slide(prs)

    add_section_title(slide, MARGIN_L, MARGIN_T, Inches(7),
                      "Part II. 멀티모달 생리반응 모델링",
                      "")

    add_rect(slide, Inches(8.0), Inches(0.5), Inches(5.0), Inches(0.55), fill_color=LIGHT_BG)
    add_multiline_textbox(
        slide, Inches(8.15), Inches(0.55), Inches(4.7), Inches(0.4),
        [
            ("SCS_2023 (IF 11.7) + SR_2022 + IJERPH_2024", True, NAVY, 9),
            ("3편 연속 출판 시리즈", False, GRAY, 9),
        ],
        line_spacing=1.2,
    )

    # --- Experiment protocol description ---
    proto_top = Inches(1.2)
    add_rect(slide, Inches(0.6), proto_top, Inches(12.1), Inches(1.3), fill_color=LIGHT_BG)

    add_multiline_textbox(
        slide, Inches(0.8), proto_top + Inches(0.08), Inches(5.5), Inches(1.15),
        [
            ("실험 규모: 60명 × 9환경 × 2일", True, NAVY, 12),
            ("MAT (Montreal Arithmetic Task) 스트레스 유도 프로토콜", False, DARK_TEXT, 10),
            ("Day1: 준비→스트레스→자극→HRV+EEG (×6 = 60min)", False, DARK_TEXT, 10),
            ("Day2: 준비→자극→주관평가 (×6 = 42min)", False, DARK_TEXT, 10),
        ],
        line_spacing=1.25,
    )

    add_multiline_textbox(
        slide, Inches(6.5), proto_top + Inches(0.08), Inches(6.0), Inches(1.15),
        [
            ("측정 장비:", True, NAVY, 12),
            ("HRV: SA-3000NEW (5개 시간/주파수 도메인 지표)", False, DARK_TEXT, 10),
            ("EEG: EMOTIV EPOC Flex 32ch, 128Hz sampling", False, DARK_TEXT, 10),
            ("Eye-tracking: Tobii Pro (시선 고정, 산동 분석)", False, DARK_TEXT, 10),
        ],
        line_spacing=1.25,
    )

    # --- 4 metric cards ---
    mc_top = Inches(2.7)
    mc_w = Inches(2.85)
    mc_h = Inches(0.95)
    metrics = [
        ("CCA 0.80", "물리음향 ↔ 심리반응"),
        ("CCA 0.78", "지각품질 ↔ 심리반응"),
        ("SDNN +14.6%", "스트레스 저항력 ↑"),
        ("TSI -9.5%", "스트레스 지수 ↓"),
    ]
    for i, (num, lbl) in enumerate(metrics):
        mx = Inches(0.6) + Emu(i * (int(mc_w) + int(Inches(0.2))))
        add_metric_card(slide, mx, mc_top, mc_w, mc_h, num, lbl,
                        number_size=22, label_size=10)

    # --- Two figures side by side ---
    fig_top = Inches(3.85)
    fig_h = Inches(2.4)
    add_textbox(slide, Inches(0.6), fig_top - Inches(0.25), Inches(5.8), Inches(0.25),
                "HRV Stress Recovery Analysis", font_size=10, font_color=GRAY, bold=True)
    add_image_safe(slide, img("scs2023_p9.png"),
                   Inches(0.6), fig_top, Inches(5.8), fig_h,
                   placeholder_text="scs2023_p9.png (HRV)")

    add_textbox(slide, Inches(6.7), fig_top - Inches(0.25), Inches(5.8), Inches(0.25),
                "Eye-tracking Heatmap Analysis", font_size=10, font_color=GRAY, bold=True)
    add_image_safe(slide, img("ijerph2024_p5.png"),
                   Inches(6.7), fig_top, Inches(5.8), fig_h,
                   placeholder_text="ijerph2024_p5.png (Eye)")

    # --- Implication box (full width, 3 specific points) ---
    imp_top = Inches(6.4)
    add_implication_box(slide, Inches(0.6), imp_top, Inches(12.1), Inches(0.9),
                        [
                            ("Samsung 시사점:  ", True, WHITE, 11),
                            ("(1) 생리신호 기반 객관적 품질 평가 → 주관 설문의 한계 보완  "
                             "(2) EEG·HRV로 렌더링 품질 변화에 대한 신체 반응 정량화  "
                             "(3) 사용자 몰입도를 Galaxy Watch·Buds 센서로 실시간 검증하는 파이프라인 구축 가능", False, WHITE, 9),
                        ],
                        font_size=10, line_spacing=1.3)


# ============================================================================
# S13: Soundscape Design Application
# ============================================================================
def build_s13(prs):
    slide = add_blank_slide(prs)

    add_section_title(slide, MARGIN_L, MARGIN_T, Inches(7),
                      "사운드스케이프 디자인 응용",
                      "")

    add_rect(slide, Inches(8.0), Inches(0.5), Inches(5.0), Inches(0.55), fill_color=LIGHT_BG)
    add_multiline_textbox(
        slide, Inches(8.15), Inches(0.55), Inches(4.7), Inches(0.4),
        [
            ("B&E_2021 + B&E_2022", True, NAVY, 11),
            ("Building and Environment, IF 7.4 x 2편", False, GRAY, 9),
        ],
        line_spacing=1.2,
    )

    # RQ
    rq_top = Inches(1.25)
    add_rect(slide, Inches(0.6), rq_top, Inches(12.1), Inches(0.65), fill_color=LIGHT_BG)
    add_textbox(slide, Inches(0.8), rq_top + Inches(0.08), Inches(11.7), Inches(0.5),
                "RQ: 오디오-비주얼 상호작용이 실내 환경의 업무 품질과 생산성에 미치는 영향은?",
                font_size=12, font_color=NAVY, bold=True)

    # === LEFT: Two paper summaries ===
    left_x = Inches(0.6)
    meth_top = Inches(2.1)
    meth_w = Inches(5.8)

    add_multiline_textbox(
        slide, left_x, meth_top, meth_w, Inches(2.0),
        [
            ("B&E 2021 — SEM 구조방정식 모델", True, NAVY, 12),
            ("  Audio → Visual → 환경만족도 경로계수 정량화", False, DARK_TEXT, 10),
            ("  오디오 품질이 시각 쾌적성에 간접효과 (cross-modal)", False, DARK_TEXT, 10),
            ("  Soundscape→Overall satisfaction 경로 유의 (p<0.01)", False, DARK_TEXT, 10),
            ("", False, DARK_TEXT, 6),
            ("B&E 2022 — 실내 사운드스케이프 → 업무 품질", True, NAVY, 12),
            ("  Audio-Visual 일치 콘텐츠가 업무 선호도·생산성 유의 향상", False, DARK_TEXT, 10),
            ("  자연 사운드스케이프: 집중도 +15%, 스트레스 -12%", False, DARK_TEXT, 10),
            ("  도시 소음: 업무 정확도 -8% (방해 효과 정량화)", False, DARK_TEXT, 10),
        ],
        line_spacing=1.25,
    )

    # === RIGHT: SEM figure (LARGE) ===
    fig_x = Inches(6.7)
    fig_top = Inches(2.1)
    add_textbox(slide, fig_x, fig_top, Inches(6.0), Inches(0.3),
                "SEM Path Model — Path Coefficients", font_size=11, font_color=GRAY, bold=True)
    add_image_safe(slide, img("be2021_p9.png"),
                   fig_x, fig_top + Inches(0.35), Inches(6.0), Inches(3.5),
                   placeholder_text="be2021_p9.png (SEM)")

    # === BOTTOM: Metric cards + Implication ===
    mc_top = Inches(5.8)
    mc_h = Inches(0.8)
    mc_w = Inches(2.5)
    metrics = [
        ("+15%", "자연 사운드 집중도 향상"),
        ("-12%", "스트레스 감소 효과"),
        ("p < 0.01", "SEM 경로 유의성"),
    ]
    for i, (num, lbl) in enumerate(metrics):
        mx = Inches(0.6) + Emu(i * (int(mc_w) + int(Inches(0.2))))
        add_metric_card(slide, mx, mc_top, mc_w, mc_h, num, lbl,
                        number_size=22, label_size=9)

    add_implication_box(slide, Inches(8.4), mc_top, Inches(4.3), Inches(1.5),
                        [
                            ("Samsung 시사점", True, WHITE, 12),
                            ("Eclipsa Audio 재생 공간의 실내 환경 설계:", False, WHITE, 10),
                            ("→ TV·사운드바 + 조명의 A-V 통합 최적화", False, WHITE, 10),
                            ("→ Galaxy Home 환경 음향 자동 튜닝 근거", False, WHITE, 10),
                        ],
                        font_size=10, line_spacing=1.25)


# ============================================================================
# S14: AVAS Soundscape Design
# ============================================================================
def build_s14(prs):
    slide = add_blank_slide(prs)

    add_section_title(slide, MARGIN_L, MARGIN_T, Inches(7),
                      "Part III. AVAS 사운드스케이프 디자인",
                      "")

    add_rect(slide, Inches(8.0), Inches(0.5), Inches(5.0), Inches(0.55), fill_color=LIGHT_BG)
    add_multiline_textbox(
        slide, Inches(8.15), Inches(0.55), Inches(4.7), Inches(0.4),
        [
            ("HMG 학술대회 특별상 + JASA (심사 중)", True, NAVY, 10),
            ("134명 대규모 청감평가", False, GRAY, 9),
        ],
        line_spacing=1.2,
    )

    # Data overview bar
    data_top = Inches(1.2)
    add_rect(slide, Inches(0.6), data_top, Inches(12.1), Inches(0.5), fill_color=LIGHT_BG)
    add_textbox(slide, Inches(0.8), data_top + Inches(0.05), Inches(11.7), Inches(0.4),
                "17개 EV × 43개 AVAS 바이노럴 녹음 | 134명 청감평가 | Binaural BHS II + GoPro Visual | 34대 경쟁 DB",
                font_size=11, font_color=DARK_TEXT, bold=True, alignment=PP_ALIGN.CENTER)

    # === LEFT: Methodology + Results ===
    left_x = Inches(0.6)
    meth_top = Inches(1.9)
    meth_w = Inches(5.8)

    add_multiline_textbox(
        slide, left_x, meth_top, meth_w, Inches(2.5),
        [
            ("3단계 감성어휘 개발:", True, NAVY, 12),
            ("  Stage 1: 272개 형용사 수집", False, DARK_TEXT, 10),
            ("  Stage 2: 전문가 축소 → 25쌍", False, DARK_TEXT, 10),
            ("  Stage 3: 요인분석 최종 18쌍", False, DARK_TEXT, 10),
            ("", False, DARK_TEXT, 6),
            ("PCA 결과 — 신규 평가축:", True, ACCENT_BLUE, 12),
            ("  기존: Pleasant-Eventful (환경 사운드스케이프)", False, DARK_TEXT, 10),
            ("  신규: Comfort-Metallic (EV AVAS 특화 축)", False, DARK_TEXT, 10),
            ("  → 기존 프레임워크로 설명 불가한 차량 음질 차원 발견", False, DARK_TEXT, 10),
            ("", False, DARK_TEXT, 6),
            ("브랜드 포지셔닝:", True, ACCENT_BLUE, 12),
            ("  34대 경쟁 차량 DB로 브랜드별 AVAS 음질 맵핑", False, DARK_TEXT, 10),
            ("  만족도 예측 모델: 정확도 92.5%", False, DARK_TEXT, 10),
        ],
        line_spacing=1.2,
    )

    # === RIGHT: PCA figure (LARGE) ===
    fig_x = Inches(6.7)
    fig_top = Inches(1.9)
    add_textbox(slide, fig_x, fig_top, Inches(6.0), Inches(0.3),
                "PCA — Comfort-Metallic Axes + Brand Positioning", font_size=10, font_color=GRAY, bold=True)
    add_image_safe(slide, img("hmg_s12_img1.png"),
                   fig_x, fig_top + Inches(0.35), Inches(6.0), Inches(3.3),
                   placeholder_text="PCA Figure: Comfort-Metallic Axes")

    # === BOTTOM: Metric cards + Implication ===
    mc_top = Inches(5.7)
    mc_h = Inches(0.8)
    mc_w = Inches(2.0)
    metrics = [
        ("92.5%", "만족도 예측 정확도"),
        ("18쌍", "최종 감성어휘"),
        ("34대", "경쟁 DB 벤치마크"),
    ]
    for i, (num, lbl) in enumerate(metrics):
        mx = Inches(0.6) + Emu(i * (int(mc_w) + int(Inches(0.2))))
        add_metric_card(slide, mx, mc_top, mc_w, mc_h, num, lbl,
                        number_size=24, label_size=9)

    add_implication_box(slide, Inches(7.2), mc_top, Inches(5.5), Inches(1.55),
                        [
                            ("Samsung 시사점", True, WHITE, 12),
                            ("대규모 청감 평가 프레임 → Eclipsa Audio 품질 인증 적용", False, WHITE, 10),
                            ("감성어휘 기반 평가축 → 렌더링 품질 차원 확장", False, WHITE, 10),
                            ("경쟁 DB 벤치마킹 → 삼성 오디오 제품 포지셔닝", False, WHITE, 10),
                            ("92.5% 예측 모델 → 렌더링 품질 튜닝 자동화 근거", False, WHITE, 10),
                        ],
                        font_size=10, line_spacing=1.2)


# ============================================================================
# S15: Research → Production Execution
# ============================================================================
def build_s15(prs):
    slide = add_blank_slide(prs)

    add_section_title(slide, MARGIN_L, MARGIN_T, Inches(12),
                      "연구 → 양산 실행력",
                      "AVAS 브랜드 사운드 2.0 — EV3, IONIQ 5 양산 적용")

    # --- Visual process flow: 4 large boxes with arrows ---
    steps = [
        ("음향 설계", "소리 특성 정의\n목표 음질 설정\n감성어휘 기반 디자인", LIGHT_BG, NAVY),
        ("청취 평가", "134명 대규모 실험\n18쌍 감성어휘\nPCA 만족도 모델", LIGHT_BG, NAVY),
        ("시스템 검증", "실차 바이노럴 녹음\nBHS II + GoPro\n주행 조건별 검증", LIGHT_BG, NAVY),
        ("양산 적용", "EV3 양산 적용\nIONIQ 5 적용\nShiny 웹 예측 툴", NAVY, WHITE),
    ]
    box_w = Inches(2.65)
    box_h = Inches(1.6)
    flow_top = Inches(1.55)
    for i, (title, detail, bg, fc) in enumerate(steps):
        bx = Inches(0.5) + Inches(i * 3.2)
        add_rect(slide, bx, flow_top, box_w, box_h, fill_color=bg)
        # Step number circle
        add_rect(slide, bx + Inches(0.1), flow_top + Inches(0.08), Inches(0.3), Inches(0.3),
                 fill_color=ACCENT_BLUE if bg != NAVY else WHITE)
        add_textbox(slide, bx + Inches(0.1), flow_top + Inches(0.08), Inches(0.3), Inches(0.3),
                    str(i + 1), font_size=12, font_color=WHITE if bg != NAVY else NAVY, bold=True,
                    alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE, font_name=FONT_EN)
        # Title
        add_textbox(slide, bx + Inches(0.5), flow_top + Inches(0.08), box_w - Inches(0.6), Inches(0.3),
                    title, font_size=13, font_color=fc, bold=True)
        # Detail
        add_textbox(slide, bx + Inches(0.15), flow_top + Inches(0.5), box_w - Inches(0.3), Inches(1.0),
                    detail, font_size=10, font_color=fc if fc == WHITE else DARK_TEXT)
        # Arrow
        if i < 3:
            add_textbox(slide, bx + box_w, flow_top + Inches(0.55), Inches(0.55), Inches(0.5),
                        "→", font_size=20, font_color=ACCENT_BLUE, bold=True,
                        alignment=PP_ALIGN.CENTER, font_name=FONT_EN)

    # --- 4 achievement cards ---
    ac_top = Inches(3.4)
    ac_w = Inches(2.85)
    ac_h = Inches(0.95)
    achievements = [
        ("특허 6건", "국내 4건 + 미국 2건"),
        ("기술이전 5천만원", "산업 적용 → 양산 실현"),
        ("HMG 특별상", "학술대회 연구 성과상"),
        ("Shiny 웹 툴", "실시간 만족도 예측"),
    ]
    for i, (num, lbl) in enumerate(achievements):
        ax = Inches(0.5) + Emu(i * (int(ac_w) + int(Inches(0.2))))
        add_metric_card(slide, ax, ac_top, ac_w, ac_h, num, lbl,
                        number_size=20, label_size=9)

    # --- Cross-department coordination ---
    coord_top = Inches(4.6)
    add_rect(slide, Inches(0.5), coord_top, Inches(12.2), Inches(0.5), fill_color=LIGHT_BG)
    add_textbox(slide, Inches(0.7), coord_top + Inches(0.05), Inches(11.8), Inches(0.4),
                "부서 간 협업: Design (사운드 아이덴티티) × Regulation (UN R138 법규 대응) × NVH (실차 검증) × 양산팀 (적용)",
                font_size=10, font_color=DARK_TEXT, bold=True, alignment=PP_ALIGN.CENTER)

    # --- Figure + Implication ---
    add_image_safe(slide, img("hmg_s12_img3.png"),
                   Inches(0.5), Inches(5.3), Inches(6.5), Inches(2.0),
                   placeholder_text="Production Process Figure")

    add_implication_box(slide, Inches(7.3), Inches(5.3), Inches(5.4), Inches(2.0),
                        [
                            ("Samsung 시사점", True, WHITE, 12),
                            ("", False, WHITE, 4),
                            ("연구→양산 end-to-end 실행력 입증:", False, WHITE, 10),
                            ("  청감평가 → 특허 → 기술이전 → 양산의 전 주기", False, WHITE, 10),
                            ("Samsung 적용:", True, WHITE, 10),
                            ("  렌더링 알고리즘의 제품 적용 주기 단축", False, WHITE, 10),
                            ("  Eclipsa Audio 품질 인증 → 제품 출시 파이프라인", False, WHITE, 10),
                        ],
                        font_size=10, line_spacing=1.2)


# ============================================================================
# S16: AI Audio Processing
# ============================================================================
def build_s16(prs):
    slide = add_blank_slide(prs)

    add_section_title(slide, MARGIN_L, MARGIN_T, Inches(7),
                      "Part IV. AI 기반 오디오 처리",
                      "")

    add_rect(slide, Inches(8.0), Inches(0.5), Inches(5.0), Inches(0.55), fill_color=LIGHT_BG)
    add_multiline_textbox(
        slide, Inches(8.15), Inches(0.55), Inches(4.7), Inches(0.4),
        [
            ("SENSORS_2021 + 특허 2건", True, NAVY, 11),
            ("기술이전 5천만원 달성", False, GRAY, 9),
        ],
        line_spacing=1.2,
    )

    # === LEFT TOP: Problem & Solution ===
    left_x = Inches(0.6)
    ps_top = Inches(1.2)
    ps_w = Inches(6.0)

    add_rect(slide, left_x, ps_top, ps_w, Inches(1.6), fill_color=LIGHT_BG)
    add_multiline_textbox(
        slide, left_x + Inches(0.15), ps_top + Inches(0.1), ps_w - Inches(0.3), Inches(1.4),
        [
            ("Problem: 데이터 희소성", True, RGBColor(0xC0, 0x39, 0x2B), 12),
            ("  30명 전문가 × 126건 청감평가 → 심각한 데이터 부족", False, DARK_TEXT, 10),
            ("  전문가 청감평가: 비용·시간 과다, 재현성 한계", False, DARK_TEXT, 10),
            ("", False, DARK_TEXT, 4),
            ("Solution: RIR Convolution Data Augmentation", True, RGBColor(0x27, 0xAE, 0x60), 12),
            ("  8개 원본 → RIR 합성곱 → 43,000개로 확장 (5,375배)", False, DARK_TEXT, 10),
            ("  + 신규 Feature: Loudness ISO 532B + Energy Ratio", False, DARK_TEXT, 10),
        ],
        line_spacing=1.25,
    )

    # === RIGHT TOP: LSTM figure ===
    fig_x = Inches(6.9)
    fig_top = Inches(1.2)
    add_textbox(slide, fig_x, fig_top, Inches(5.8), Inches(0.25),
                "5-Layer LSTM Architecture", font_size=10, font_color=GRAY, bold=True)
    add_image_safe(slide, img("sensors2021_p5.png"),
                   fig_x, fig_top + Inches(0.3), Inches(5.8), Inches(2.3),
                   placeholder_text="sensors2021_p5.png (LSTM)")

    # === MIDDLE: AI vs 8 Pulmonologists comparison ===
    comp_top = Inches(3.1)
    add_textbox(slide, left_x, comp_top, Inches(6.0), Inches(0.3),
                "AI vs 8명 호흡기내과 전문의 비교", font_size=12, font_color=NAVY, bold=True)

    comparisons = [
        ("Accuracy", "84.9%", "56.4%"),
        ("Sensitivity", "90.0%", "40.7%"),
        ("AUC", "0.84", "0.56"),
    ]
    cc_w = Inches(2.6)
    cc_h = Inches(1.2)
    cc_top = Inches(3.45)
    for i, (title, ai_val, expert_val) in enumerate(comparisons):
        cx = Inches(0.6) + Emu(i * (int(cc_w) + int(Inches(0.2))))
        add_rect(slide, cx, cc_top, cc_w, cc_h, fill_color=LIGHT_BG,
                 line_color=ACCENT_BLUE, line_width_pt=1.0)
        add_textbox(slide, cx, cc_top + Inches(0.05), cc_w, Inches(0.25),
                    title, font_size=10, font_color=GRAY, bold=True,
                    alignment=PP_ALIGN.CENTER, font_name=FONT_EN)
        add_textbox(slide, cx, cc_top + Inches(0.3), cc_w, Inches(0.4),
                    f"AI  {ai_val}", font_size=20, font_color=NAVY, bold=True,
                    alignment=PP_ALIGN.CENTER, font_name=FONT_EN)
        add_textbox(slide, cx, cc_top + Inches(0.75), cc_w, Inches(0.35),
                    f"Expert avg  {expert_val}", font_size=11, font_color=GRAY,
                    alignment=PP_ALIGN.CENTER, font_name=FONT_EN)

    # === IP box (right middle) ===
    ip_top = Inches(3.7)
    ip_w = Inches(4.6)
    add_rect(slide, Inches(8.2), ip_top, ip_w, Inches(1.0), fill_color=LIGHT_BG,
             line_color=NAVY, line_width_pt=1.0)
    add_multiline_textbox(
        slide, Inches(8.35), ip_top + Inches(0.08), ip_w - Inches(0.3), Inches(0.8),
        [
            ("IP & 기술이전", True, NAVY, 12),
            ("KR 특허 등록 + US 특허 출원", False, DARK_TEXT, 10),
            ("기술이전: 5,000만원 (현대자동차 → 산업체)", False, DARK_TEXT, 10),
        ],
        line_spacing=1.3,
    )

    # === A-JEPA Vision box (PROMINENT, bottom left) ===
    ajepa_top = Inches(4.9)
    ajepa_w = Inches(7.5)
    ajepa_h = Inches(1.15)
    add_rect(slide, Inches(0.6), ajepa_top, ajepa_w, ajepa_h, fill_color=ACCENT_BLUE)
    add_multiline_textbox(
        slide, Inches(0.8), ajepa_top + Inches(0.1), ajepa_w - Inches(0.4), ajepa_h - Inches(0.2),
        [
            ("A-JEPA Vision — 차세대 AI Audio 연구 방향", True, WHITE, 14),
            ("Meta Audio-JEPA 자기지도 학습 + 음향 도메인 지식 결합", False, WHITE, 11),
            ("→ 라벨 없는 대규모 오디오 데이터에서 음질 표현 자동 학습", False, WHITE, 11),
            ("→ Eclipsa Audio 렌더링 품질의 end-to-end 자동 판정 목표", False, WHITE, 11),
        ],
        font_color=WHITE,
        line_spacing=1.2,
    )

    # === Bottom right: Implication ===
    add_implication_box(slide, Inches(8.4), ajepa_top, Inches(4.3), Inches(2.3),
                        [
                            ("Samsung 시사점", True, WHITE, 12),
                            ("", False, WHITE, 4),
                            ("Eclipsa Audio 렌더링 품질의", False, WHITE, 10),
                            ("자동 판정 파이프라인 구축", False, WHITE, 10),
                            ("", False, WHITE, 4),
                            ("대규모 A/B 테스트 비용 절감", False, WHITE, 10),
                            ("(전문가 8명 수준 → AI 1개 모델)", False, WHITE, 10),
                            ("", False, WHITE, 4),
                            ("A-JEPA로 비지도 학습 확장", False, WHITE, 10),
                        ],
                        font_size=10, line_spacing=1.1)


# ============================================================================
# S17: Contribution Plan
# ============================================================================
def build_s17(prs):
    slide = add_blank_slide(prs)

    add_section_title(slide, MARGIN_L, MARGIN_T, Inches(12),
                      "Contribution Plan",
                      "삼성리서치 Spatial Audio 기여 로드맵")

    # 3 timeline column headers
    cols = [
        ("입사 ~ 6개월 (즉시 기여)", ACCENT_BLUE),
        ("6개월 ~ 2년 (과제 확장)", NAVY),
        ("2년 ~ (Lab 비전 주도)", RGBColor(0x0A, 0x1A, 0x6E)),
    ]
    col_w = Inches(3.45)
    col_h = Inches(0.45)
    label_w = Inches(1.9)
    hdr_top = Inches(1.55)

    for i, (title, color) in enumerate(cols):
        cx = label_w + Inches(0.15) + Emu(i * (int(col_w) + int(Inches(0.12))))
        add_rect(slide, cx, hdr_top, col_w, col_h, fill_color=color)
        add_textbox(slide, cx, hdr_top, col_w, col_h,
                    title, font_size=10, font_color=WHITE, bold=True,
                    alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # 4 row labels + cells
    row_labels = ["Eclipsa\nAudio", "AI 전환", "파트\n시너지", "인증·표준"]
    row_h = Inches(1.1)
    row_start = hdr_top + col_h + Inches(0.08)

    grid = [
        [
            "청감 평가 프레임 구축\n(ISO 12913 + SATP 방법론)\n→ 기존 연구 즉시 적용 가능",
            "HRTF 개인화 알고리즘 개발\n렌더러 품질 최적화\n→ 77% 공간감 기여도 활용",
            "차세대 Eclipsa Audio\n품질 표준 주도\n→ 글로벌 de facto 표준 목표",
        ],
        [
            "A-JEPA 프로토타입 구축\n음질 자동 판정 파이프라인\n→ LSTM 84.9% 정확도 기반",
            "생리신호 기반\n실시간 품질 모니터링\n→ Galaxy Watch 연동",
            "AI 청감 평가 플랫폼\n자동화 완성\n→ 비지도 학습 확장",
        ],
        [
            "Display팀 협업 시작\nA-V 통합 평가 설계\n→ Visual 76% 기여도 활용",
            "Holographic Display ×\nSpatial Audio 시너지\n→ cross-modal 효과 극대화",
            "크로스모달 경험 설계\nLab 비전 제안\n→ 차세대 제품 전략",
        ],
        [
            "Eclipsa Audio 인증 기준\n초안 작성\n→ 92.5% 예측모델 활용",
            "국제 표준화 기여\n(IEC/ISO WG 참여)\n→ SATP 18개국 네트워크",
            "삼성 주도 표준\nde facto 확립\n→ 산업 리더십 확보",
        ],
    ]

    for r, (label, cells) in enumerate(zip(row_labels, grid)):
        ry = row_start + Emu(r * (int(row_h) + int(Inches(0.08))))
        # Row label
        add_rect(slide, Inches(0.15), ry, label_w, row_h, fill_color=LIGHT_BG)
        add_textbox(slide, Inches(0.15), ry, label_w, row_h,
                    label, font_size=10, font_color=NAVY, bold=True,
                    alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # Cells
        for c, cell_text in enumerate(cells):
            cx = label_w + Inches(0.15) + Emu(c * (int(col_w) + int(Inches(0.12))))
            add_rect(slide, cx, ry, col_w, row_h, fill_color=LIGHT_BG,
                     line_color=WARM_GRAY, line_width_pt=0.5)
            add_textbox(slide, cx + Inches(0.1), ry + Inches(0.05),
                        col_w - Inches(0.2), row_h - Inches(0.1),
                        cell_text, font_size=9, font_color=DARK_TEXT,
                        alignment=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE)

    # Bottom navy bar with tagline
    bar_top = Inches(6.55)
    add_rect(slide, Inches(0.15), bar_top, Inches(13.0), Inches(0.65), fill_color=NAVY)
    add_textbox(slide, Inches(0.35), bar_top + Inches(0.1), Inches(12.6), Inches(0.45),
                "신호처리가 만드는 기술 위에, 사용자가 경험하는 품질을 설계합니다 "
                "— Perception-driven Spatial Audio Quality Engineering",
                font_size=12, font_color=WHITE, bold=True,
                alignment=PP_ALIGN.CENTER)


# ============================================================================
# S18: Thank You
# ============================================================================
def build_s18(prs):
    slide = add_blank_slide(prs)
    set_slide_bg(slide, NAVY)

    add_textbox(slide, Inches(0.8), Inches(2.0), Inches(11.5), Inches(1.0),
                "THANK YOU",
                font_size=36, font_color=WHITE, bold=True,
                alignment=PP_ALIGN.CENTER, font_name=FONT_EN)

    add_textbox(slide, Inches(0.8), Inches(3.2), Inches(11.5), Inches(0.5),
                "경청해 주셔서 감사합니다. 질문 부탁드립니다.",
                font_size=14, font_color=WARM_GRAY, bold=False,
                alignment=PP_ALIGN.CENTER)

    add_accent_bar(slide, Inches(5.0), Inches(4.0), Inches(3.3),
                   Inches(0.04), color=ACCENT_BLUE)

    add_textbox(slide, Inches(0.8), Inches(4.5), Inches(11.5), Inches(0.4),
                "조현인 (Hyun In Jo, Ph.D.)",
                font_size=13, font_color=WHITE, bold=True,
                alignment=PP_ALIGN.CENTER)

    add_textbox(slide, Inches(0.8), Inches(5.0), Inches(11.5), Inches(0.4),
                "best2012@naver.com | 010-6387-8402 | linkedin.com/in/hyunin-jo",
                font_size=10, font_color=WARM_GRAY, bold=False,
                alignment=PP_ALIGN.CENTER, font_name=FONT_EN)

    add_textbox(slide, Inches(0.8), Inches(5.5), Inches(11.5), Inches(0.4),
                "Samsung Research · Visual Technology · Display Innovation Lab · Spatial Audio",
                font_size=10, font_color=WARM_GRAY, bold=False,
                alignment=PP_ALIGN.CENTER)


# ============================================================================
# Main
# ============================================================================
def main():
    prs = new_presentation()

    builders = [
        build_s1, build_s2, build_s3, build_s4, build_s5, build_s6,
        build_s7, build_s8, build_s9, build_s10, build_s11, build_s12,
        build_s13, build_s14, build_s15, build_s16, build_s17, build_s18,
    ]

    for fn in builders:
        fn(prs)

    out_path = os.path.join(os.path.dirname(os.path.dirname(__file__)),
                            "Portfolio_HyunInJo_v2.pptx")
    prs.save(out_path)
    print(f"Saved: {out_path} ({len(prs.slides)} slides)")


if __name__ == "__main__":
    main()
