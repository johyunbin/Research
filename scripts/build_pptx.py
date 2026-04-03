#!/usr/bin/env python3
"""
build_pptx.py — Build all 18 slides of the Samsung Research portfolio PPT.
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
        # Dot
        add_rect(slide, tl_left, y + Inches(0.08), Inches(0.12), Inches(0.12),
                 fill_color=NAVY)
        # Connecting line (except last)
        if i < len(entries) - 1:
            add_rect(slide, tl_left + Inches(0.05), y + Inches(0.2),
                     Inches(0.02), Inches(0.5), fill_color=ACCENT_BLUE)
        # Year
        add_textbox(slide, tl_left + Inches(0.3), y, Inches(1.5), Inches(0.3),
                    year, font_size=10, font_color=ACCENT_BLUE, bold=True,
                    font_name=FONT_EN)
        # Description
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

    # Top text
    add_textbox(slide, Inches(0.8), Inches(0.5), Inches(11.5), Inches(1.0),
                "THD 0.01%, 주파수 응답 ±0.5dB —\n공학 스펙이 완벽해도 사용자가 \"좋다\"고 느끼지 않을 수 있다",
                font_size=16, font_color=GRAY, bold=False,
                alignment=PP_ALIGN.CENTER)

    # Mid-top accent
    add_textbox(slide, Inches(0.8), Inches(1.6), Inches(11.5), Inches(0.5),
                "시각 맥락만으로 오디오 만족도가 76% 좌우된다면?",
                font_size=14, font_color=ACCENT_BLUE, bold=True,
                alignment=PP_ALIGN.CENTER)

    # Center navy rectangle
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

    # Bottom: 4 Part roadmap cards
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
        # Arrow between cards
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

    # Left box (Light BG)
    lw = Inches(5.0)
    lh = Inches(1.5)
    lt = Inches(1.8)
    add_rect(slide, Inches(0.6), lt, lw, lh, fill_color=LIGHT_BG)
    add_textbox(slide, Inches(0.8), lt + Inches(0.15), lw - Inches(0.4), Inches(0.3),
                "Traditional Noise Control", font_size=13, font_color=NAVY, bold=True)
    add_textbox(slide, Inches(0.8), lt + Inches(0.55), lw - Inches(0.4), Inches(0.7),
                '"소음이 얼마나 큰가?" dB 측정',
                font_size=11, font_color=DARK_TEXT)

    # Arrow
    add_textbox(slide, Inches(5.8), lt + Inches(0.4), Inches(1.2), Inches(0.5),
                "→", font_size=28, font_color=ACCENT_BLUE, bold=True,
                alignment=PP_ALIGN.CENTER, font_name=FONT_EN)

    # Right box (Navy)
    add_rect(slide, Inches(7.2), lt, lw, lh, fill_color=NAVY)
    add_textbox(slide, Inches(7.4), lt + Inches(0.15), lw - Inches(0.4), Inches(0.3),
                "New Paradigm: Soundscape", font_size=13, font_color=WHITE, bold=True)
    add_textbox(slide, Inches(7.4), lt + Inches(0.55), lw - Inches(0.4), Inches(0.7),
                '"소리가 어떻게 경험되는가?" 인간 지각 중심',
                font_size=11, font_color=WHITE)

    # ISO 12913 definition
    add_rect(slide, Inches(0.6), Inches(3.6), Inches(11.8), Inches(0.8), fill_color=LIGHT_BG)
    add_textbox(slide, Inches(0.8), Inches(3.7), Inches(11.4), Inches(0.6),
                "ISO 12913: \"acoustic environment as perceived or experienced and/or understood "
                "by a person or people, in context\"",
                font_size=11, font_color=GRAY, bold=False,
                alignment=PP_ALIGN.CENTER)

    # Bottom: Pleasant-Eventful diagram
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
        # Left cell
        add_rect(slide, left1, ry, col_w, row_h, fill_color=LIGHT_BG)
        add_textbox(slide, left1 + Inches(0.15), ry + Inches(0.08),
                    col_w - Inches(0.3), row_h - Inches(0.16),
                    left_t, font_size=10, font_color=DARK_TEXT)
        # Arrow
        add_textbox(slide, arrow_l, ry + Inches(0.15), Inches(1.0), Inches(0.5),
                    "→", font_size=16, font_color=ACCENT_BLUE, bold=True,
                    alignment=PP_ALIGN.CENTER, font_name=FONT_EN)
        # Right cell
        add_rect(slide, left2, ry, col_w, row_h, fill_color=LIGHT_BG)
        add_textbox(slide, left2 + Inches(0.15), ry + Inches(0.08),
                    col_w - Inches(0.3), row_h - Inches(0.16),
                    right_t, font_size=10, font_color=DARK_TEXT)

    # Bottom navy bar with tagline
    bar_top = Inches(6.5)
    add_rect(slide, Inches(0.6), bar_top, Inches(12.1), Inches(0.7), fill_color=NAVY)
    add_textbox(slide, Inches(0.8), bar_top + Inches(0.1), Inches(11.7), Inches(0.5),
                '신호처리가 "어떻게 구현할 것인가"라면, 저의 전문성은 '
                '"사용자가 어떻게 경험할 것인가"를 설계하고 검증하는 것입니다',
                font_size=11, font_color=WHITE, bold=True,
                alignment=PP_ALIGN.CENTER)


# ============================================================================
# S6: Spatial Audio Overview
# ============================================================================
def build_s6(prs):
    slide = add_blank_slide(prs)

    add_section_title(slide, MARGIN_L, MARGIN_T, Inches(12),
                      "Part I. Spatial Audio 기술 역량",
                      "5개 기술 축 연구 체계도")

    # Center navy bar
    bar_w = Inches(8.0)
    bar_l = Inches((13.333 - 8.0) / 2)
    add_rect(slide, bar_l, Inches(2.0), bar_w, Inches(0.6), fill_color=NAVY)
    add_textbox(slide, bar_l, Inches(2.0), bar_w, Inches(0.6),
                "Spatial Audio Perception Research",
                font_size=16, font_color=WHITE, bold=True,
                alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE,
                font_name=FONT_EN)

    # 5 axis cards
    axes = [
        ("축1", "시각 재현 방식", "APAC'22"),
        ("축2", "재생 방식", "APAC'19"),
        ("축3", "바이노럴-비주얼", "B&E'19 ×2"),
        ("축4", "시청각 정보", "B&E'20"),
        ("축5", "생태학적 타당성", "SCS'21"),
    ]
    card_w = Inches(2.2)
    card_h = Inches(1.6)
    start_x = Inches(0.55)
    card_top = Inches(3.5)

    for i, (ax_label, ax_title, ax_ref) in enumerate(axes):
        cx = start_x + Inches(i * 2.5)
        # Connecting line from bar
        line_x = cx + card_w / 2 - Inches(0.01)
        add_rect(slide, line_x, Inches(2.6), Inches(0.02), Inches(0.9),
                 fill_color=ACCENT_BLUE)
        # Card
        add_rect(slide, cx, card_top, card_w, card_h, fill_color=LIGHT_BG)
        add_textbox(slide, cx, card_top + Inches(0.15), card_w, Inches(0.3),
                    ax_label, font_size=12, font_color=ACCENT_BLUE, bold=True,
                    alignment=PP_ALIGN.CENTER, font_name=FONT_EN)
        add_textbox(slide, cx + Inches(0.1), card_top + Inches(0.5),
                    card_w - Inches(0.2), Inches(0.5),
                    ax_title, font_size=13, font_color=NAVY, bold=True,
                    alignment=PP_ALIGN.CENTER)
        add_textbox(slide, cx, card_top + Inches(1.1), card_w, Inches(0.3),
                    ax_ref, font_size=10, font_color=GRAY, bold=False,
                    alignment=PP_ALIGN.CENTER, font_name=FONT_EN)


# ============================================================================
# S7-S11: Research Axis Slides (reusable template)
# ============================================================================
def _build_research_axis(prs, section_title, paper, rq, metrics, result, implication, image_file):
    """
    Reusable template for S7-S11.
    metrics: list of (number, label) tuples
    """
    slide = add_blank_slide(prs)

    add_section_title(slide, MARGIN_L, MARGIN_T, Inches(12),
                      section_title, paper, subtitle_size=11)

    # RQ box
    rq_top = Inches(1.6)
    add_rect(slide, Inches(0.6), rq_top, Inches(12.1), Inches(0.8), fill_color=LIGHT_BG)
    add_textbox(slide, Inches(0.8), rq_top + Inches(0.1), Inches(11.7), Inches(0.6),
                f"RQ: {rq}", font_size=12, font_color=NAVY, bold=True,
                alignment=PP_ALIGN.LEFT)

    # Metric cards
    n_metrics = len(metrics)
    mc_w = Inches(2.5)
    mc_h = Inches(0.85)
    mc_top = Inches(2.6)
    total_w = mc_w * n_metrics + Inches(0.2) * (n_metrics - 1)
    mc_start = Inches((13.333 - total_w / Emu(Inches(1))) / 2) if n_metrics <= 3 else Inches(0.6)
    # simpler spacing
    mc_start = Inches(0.6)
    gap = (Inches(12.1) - mc_w * n_metrics) // max(n_metrics - 1, 1) if n_metrics > 1 else 0

    for i, (num, lbl) in enumerate(metrics):
        mx = Inches(0.6) + Emu(i * (int(mc_w) + int(Inches(0.3))))
        add_metric_card(slide, mx, mc_top, mc_w, mc_h, num, lbl,
                        number_size=20, label_size=9)

    # Result line
    res_top = Inches(3.7)
    add_textbox(slide, Inches(0.6), res_top, Inches(12.1), Inches(0.5),
                f"▸ {result}", font_size=11, font_color=DARK_TEXT, bold=False)

    # Figure image (left area)
    add_image_safe(slide, img(image_file),
                   Inches(0.6), Inches(4.3), Inches(6.5), Inches(2.7),
                   placeholder_text=image_file)

    # Implication box (right side)
    add_implication_box(slide, Inches(7.4), Inches(4.3), Inches(5.3), Inches(1.2),
                        [f"▸ Samsung 시사점: {implication}"],
                        font_size=11)


def build_s7(prs):
    _build_research_axis(
        prs,
        section_title="축1: 시각 재현 방식",
        paper="APAC_2022_Jo&Jeon (Applied Acoustics, IF 3.4)",
        rq="같은 Spatial Audio를 HMD vs 모니터로 볼 때, 사용자 지각이 얼마나 달라지는가?",
        metrics=[
            ("40명 × 8환경", "실험 규모"),
            ("p < 0.05", "HMD에서 더 민감한 인식"),
        ],
        result="HMD = 공간 현실감↑ / 모니터 = 전반적 인식↑ → 시각 재현이 오디오 판단을 변화시킴",
        implication="XR 디바이스 대응 시, 시각 조건 통제가 품질 평가의 전제",
        image_file="apac2022_p6.png",
    )


def build_s8(prs):
    _build_research_axis(
        prs,
        section_title="축2: 헤드폰 vs 스피커",
        paper="APAC_2019_Jeon et al (Applied Acoustics, IF 3.4)",
        rq="헤드폰과 스피커 재생 방식이 사용자의 소리 품질 판단을 어떻게 바꾸는가?",
        metrics=[
            ("6%", "허용한계 차이"),
            ("8%", "성가심 차이"),
            ("2.2 dBA", "50% 성가심 SPL 차이"),
        ],
        result="스피커+HMD 조합이 실제 환경에 가장 근접 → 가장 민감한 반응",
        implication="헤드폰(바이노럴) vs 스피커(멀티채널) 지각 차이 정량화 기준",
        image_file="be2019a_p7.png",
    )


def build_s9(prs):
    _build_research_axis(
        prs,
        section_title="축3: 바이노럴-비주얼",
        paper="B&E_2019 × 2편 (Building and Environment, IF 7.4)",
        rq="HRTF와 HMD, 어느 것이 사용자 공간 지각에 더 지배적인가?",
        metrics=[
            ("HRTF 77%", "공간감의 지배적 요인"),
            ("HMD 23%", "시각 보조 역할"),
            ("6~7 dB↓", "VR 허용한계 하락"),
        ],
        result="HRTF+HMD 동시 적용 시 음상 외재화·몰입감 유의 증가",
        implication="HRTF 개인화 = 렌더러 최적화의 최우선 과제 (77%)",
        image_file="be2019a_p8.png",
    )


def build_s10(prs):
    _build_research_axis(
        prs,
        section_title="축4: 시청각 정보",
        paper="B&E_2020_Jeon&Jo (Building and Environment, IF 7.4)",
        rq="Audio와 Visual 정보가 전체 만족도에 각각 얼마나 기여하는가?",
        metrics=[
            ("청각 24%", "만족도 기여"),
            ("시각 76%", "만족도 기여"),
            ("설명력 51%", "만족도 모델"),
        ],
        result="오디오가 landscape 자연스러움에도 영향 (cross-modal effect)",
        implication="Display + Audio 통합 설계 → Holographic Displays 시너지",
        image_file="be2020_p8.png",
    )


def build_s11(prs):
    _build_research_axis(
        prs,
        section_title="축5: 생태학적 타당성",
        paper="SCS_2021_Jo&Jeon (Sustainable Cities and Society, IF 11.7)",
        rq="VR 실험실 평가를 실제 현장과 동일하게 신뢰할 수 있는가?",
        metrics=[
            ("50명 × 10환경", "실험 규모"),
            ("3 프로토콜", "ISO 12913-2"),
            ("VR ≈ In-situ", "유사 결과"),
        ],
        result="FOA 바이노럴 + 헤드트래킹의 높은 생태학적 타당성 실증",
        implication="VR 실험실에서 렌더링 품질 평가 → 실사용 환경 결과 신뢰 가능",
        image_file="scs2021_p5.png",
    )


# ============================================================================
# S12: Multimodal Biosignal Modeling
# ============================================================================
def build_s12(prs):
    slide = add_blank_slide(prs)

    add_section_title(slide, MARGIN_L, MARGIN_T, Inches(12),
                      "Part II. 멀티모달 생리반응 모델링",
                      "SCS_2023 (IF 11.7) + SR_2022 + IJERPH_2024")

    # Experiment info line
    add_textbox(slide, Inches(0.6), Inches(1.5), Inches(12.1), Inches(0.35),
                "60명 × 9환경 × 2일 | EEG (32ch) · HRV (5지표) · Eye-tracking",
                font_size=11, font_color=DARK_TEXT, bold=True,
                alignment=PP_ALIGN.CENTER)

    # 4 metric cards
    metrics = [
        ("CCA 0.80", "물리음향↔심리반응"),
        ("CCA 0.78", "지각품질↔심리반응"),
        ("SDNN +14.6%", "스트레스 저항력↑"),
        ("TSI -9.5%", "스트레스 지수↓"),
    ]
    mc_w = Inches(2.7)
    mc_h = Inches(0.85)
    mc_top = Inches(2.0)
    for i, (num, lbl) in enumerate(metrics):
        mx = Inches(0.6) + Emu(i * (int(mc_w) + int(Inches(0.25))))
        add_metric_card(slide, mx, mc_top, mc_w, mc_h, num, lbl,
                        number_size=18, label_size=9)

    # Two images side by side
    add_image_safe(slide, img("scs2023_p9.png"),
                   Inches(0.6), Inches(3.2), Inches(5.8), Inches(2.6),
                   placeholder_text="scs2023_p9.png")
    add_image_safe(slide, img("ijerph2024_p5.png"),
                   Inches(6.7), Inches(3.2), Inches(5.8), Inches(2.6),
                   placeholder_text="ijerph2024_p5.png")

    # Implication box
    add_implication_box(slide, Inches(0.6), Inches(6.0), Inches(12.1), Inches(1.2),
                        [
                            "▸ 생리신호 기반 객관적 품질 평가 → 주관 설문의 한계 보완",
                            "▸ EEG·HRV 지표로 렌더링 품질 변화에 대한 신체 반응 정량화",
                            "▸ Samsung: 사용자 몰입도를 생체 데이터로 실시간 검증하는 파이프라인 제안 가능",
                        ],
                        font_size=10)


# ============================================================================
# S13: Soundscape Design Application
# ============================================================================
def build_s13(prs):
    slide = add_blank_slide(prs)

    add_section_title(slide, MARGIN_L, MARGIN_T, Inches(12),
                      "사운드스케이프 디자인 응용",
                      "B&E_2021 + B&E_2022 (IF 7.4)")

    # RQ box
    rq_top = Inches(1.6)
    add_rect(slide, Inches(0.6), rq_top, Inches(12.1), Inches(0.7), fill_color=LIGHT_BG)
    add_textbox(slide, Inches(0.8), rq_top + Inches(0.1), Inches(11.7), Inches(0.5),
                "RQ: 오디오-비주얼 상호작용이 실내 환경의 업무 품질과 생산성에 미치는 영향은?",
                font_size=12, font_color=NAVY, bold=True)

    # Result text
    add_textbox(slide, Inches(0.6), Inches(2.5), Inches(12.1), Inches(0.5),
                "▸ SEM 모델로 Audio→Visual→만족도 경로 정량화 / A-V 일치 시 업무 선호도·생산성 유의 향상",
                font_size=11, font_color=DARK_TEXT)

    # SEM figure
    add_image_safe(slide, img("be2021_p9.png"),
                   Inches(0.6), Inches(3.2), Inches(7.0), Inches(3.2),
                   placeholder_text="be2021_p9.png (SEM)")

    # Implication box
    add_implication_box(slide, Inches(8.0), Inches(3.2), Inches(4.7), Inches(1.5),
                        ["▸ Samsung 시사점:",
                         "TV·사운드바 실내 환경별 최적 렌더링",
                         "설계의 이론적 근거"],
                        font_size=11)


# ============================================================================
# S14: AVAS Soundscape Design
# ============================================================================
def build_s14(prs):
    slide = add_blank_slide(prs)

    add_section_title(slide, MARGIN_L, MARGIN_T, Inches(12),
                      "Part III. AVAS 사운드스케이프 디자인",
                      "HMG 학술대회 특별상 + JASA (심사 중)")

    # Background text
    add_textbox(slide, Inches(0.6), Inches(1.5), Inches(12.1), Inches(0.35),
                "EV 시대 → AVAS가 도시 음환경의 새로운 요소",
                font_size=12, font_color=GRAY, bold=False, alignment=PP_ALIGN.CENTER)

    # Data line
    add_textbox(slide, Inches(0.6), Inches(1.9), Inches(12.1), Inches(0.35),
                "17개 EV × 43개 AVAS 바이노럴 | 134명 청감평가 | 3단계 감성어휘 (272→25→18쌍)",
                font_size=11, font_color=DARK_TEXT, bold=True, alignment=PP_ALIGN.CENTER)

    # 3 metric cards
    metrics = [
        ("Comfort–Metallic", "신규 평가축 제안"),
        ("92.5%", "만족도 예측 정확도"),
        ("34대 경쟁 DB", "벤치마킹 체계"),
    ]
    mc_w = Inches(3.5)
    mc_h = Inches(0.85)
    mc_top = Inches(2.5)
    for i, (num, lbl) in enumerate(metrics):
        mx = Inches(0.6) + Emu(i * (int(mc_w) + int(Inches(0.3))))
        add_metric_card(slide, mx, mc_top, mc_w, mc_h, num, lbl,
                        number_size=18, label_size=9)

    # Left: PCA figure
    add_image_safe(slide, img("hmg_s12_img1.png"),
                   Inches(0.6), Inches(3.6), Inches(6.5), Inches(3.2),
                   placeholder_text="PCA Figure")

    # Right: Implication box
    add_implication_box(slide, Inches(7.4), Inches(3.6), Inches(5.3), Inches(1.8),
                        [
                            "▸ Eclipsa Audio 품질 인증에 대규모 청감 평가 프레임 적용 가능",
                            "▸ 감성어휘 기반 평가축 → 렌더링 품질 차원 확장",
                            "▸ 경쟁 DB 벤치마킹 방법론 → 삼성 오디오 제품 포지셔닝",
                        ],
                        font_size=10)


# ============================================================================
# S15: Research → Production Execution
# ============================================================================
def build_s15(prs):
    slide = add_blank_slide(prs)

    add_section_title(slide, MARGIN_L, MARGIN_T, Inches(12),
                      "연구 → 양산 실행력",
                      "AVAS 브랜드 사운드 2.0 — EV3, IONIQ 5")

    # Process flow: 4 boxes with arrows
    steps = [
        ("음향 설계", LIGHT_BG, NAVY),
        ("청취 평가", LIGHT_BG, NAVY),
        ("시스템 검증", LIGHT_BG, NAVY),
        ("양산 적용", NAVY, WHITE),
    ]
    box_w = Inches(2.5)
    box_h = Inches(0.7)
    flow_top = Inches(1.6)
    for i, (txt, bg, fc) in enumerate(steps):
        bx = Inches(0.6) + Inches(i * 3.1)
        add_rect(slide, bx, flow_top, box_w, box_h, fill_color=bg)
        add_textbox(slide, bx, flow_top, box_w, box_h,
                    txt, font_size=13, font_color=fc, bold=True,
                    alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        if i < 3:
            add_textbox(slide, bx + box_w, flow_top + Inches(0.1),
                        Inches(0.6), Inches(0.5),
                        "→", font_size=18, font_color=ACCENT_BLUE, bold=True,
                        alignment=PP_ALIGN.CENTER, font_name=FONT_EN)

    # 4 achievement cards
    achievements = [
        ("특허 6건", "국내+미국"),
        ("기술이전 5천만원", "산업 적용"),
        ("HMG 특별상", "학술대회"),
        ("웹 예측 툴", "실시간 평가"),
    ]
    ac_w = Inches(2.7)
    ac_h = Inches(0.85)
    ac_top = Inches(2.6)
    for i, (num, lbl) in enumerate(achievements):
        ax = Inches(0.6) + Emu(i * (int(ac_w) + int(Inches(0.25))))
        add_metric_card(slide, ax, ac_top, ac_w, ac_h, num, lbl,
                        number_size=16, label_size=9)

    # Left: production figure
    add_image_safe(slide, img("hmg_s12_img3.png"),
                   Inches(0.6), Inches(3.8), Inches(6.5), Inches(3.0),
                   placeholder_text="Production Figure")

    # Right: Implication box
    add_implication_box(slide, Inches(7.4), Inches(3.8), Inches(5.3), Inches(1.8),
                        [
                            "▸ 연구 결과를 양산 차량에 적용한 end-to-end 실행력",
                            "▸ 청감 평가 → 특허 → 기술이전 → 양산의 전 주기 경험",
                            "▸ Samsung: 렌더링 알고리즘의 제품 적용 주기 단축 기대",
                        ],
                        font_size=10)


# ============================================================================
# S16: AI Audio Processing
# ============================================================================
def build_s16(prs):
    slide = add_blank_slide(prs)

    add_section_title(slide, MARGIN_L, MARGIN_T, Inches(12),
                      "Part IV. AI 기반 오디오 처리",
                      "SENSORS_2021 — 특허 2건, 기술이전 5천만원")

    # Problem/Solution text (left)
    add_multiline_textbox(
        slide, Inches(0.6), Inches(1.6), Inches(5.5), Inches(1.2),
        [
            ("Problem:", True, NAVY, 12),
            ("전문가 청감평가는 비용·시간 과다, 재현성 한계", False, DARK_TEXT, 11),
            ("", False, DARK_TEXT, 6),
            ("Solution:", True, ACCENT_BLUE, 12),
            ("LSTM 기반 AI 모델로 전문가 수준의 음질 판정 자동화", False, DARK_TEXT, 11),
        ],
        line_spacing=1.3,
    )

    # 3 AI vs Expert comparison cards
    comparisons = [
        ("정확도", "84.9%", "vs 56.4%"),
        ("민감도", "90.0%", "vs 40.7%"),
        ("AUC", "0.84", "vs 0.56"),
    ]
    cc_w = Inches(2.3)
    cc_h = Inches(1.0)
    cc_top = Inches(3.0)
    for i, (title, ai_val, expert) in enumerate(comparisons):
        cx = Inches(0.6) + Emu(i * (int(cc_w) + int(Inches(0.2))))
        add_rect(slide, cx, cc_top, cc_w, cc_h, fill_color=LIGHT_BG)
        add_textbox(slide, cx, cc_top + Inches(0.05), cc_w, Inches(0.25),
                    title, font_size=10, font_color=GRAY, bold=True,
                    alignment=PP_ALIGN.CENTER)
        add_textbox(slide, cx, cc_top + Inches(0.3), cc_w, Inches(0.3),
                    f"AI {ai_val}", font_size=16, font_color=NAVY, bold=True,
                    alignment=PP_ALIGN.CENTER, font_name=FONT_EN)
        add_textbox(slide, cx, cc_top + Inches(0.65), cc_w, Inches(0.25),
                    f"Expert {expert}", font_size=10, font_color=GRAY,
                    alignment=PP_ALIGN.CENTER, font_name=FONT_EN)

    # LSTM figure
    add_image_safe(slide, img("sensors2021_p5.png"),
                   Inches(7.2), Inches(1.6), Inches(5.5), Inches(2.5),
                   placeholder_text="LSTM Figure")

    # Implication box (right bottom)
    add_implication_box(slide, Inches(7.2), Inches(4.3), Inches(5.5), Inches(1.2),
                        [
                            "▸ Eclipsa Audio 렌더링 품질의 자동 판정 파이프라인",
                            "▸ 대규모 A/B 테스트 비용 절감",
                        ],
                        font_size=10)

    # A-JEPA vision box
    add_rect(slide, Inches(0.6), Inches(4.3), Inches(6.3), Inches(1.2),
             fill_color=ACCENT_BLUE)
    add_multiline_textbox(
        slide, Inches(0.8), Inches(4.4), Inches(5.9), Inches(1.0),
        [
            ("A-JEPA Vision", True, WHITE, 13),
            ("Meta A-JEPA 자기지도 학습 + 음향 도메인 지식", False, WHITE, 11),
        ],
        font_color=WHITE,
        line_spacing=1.4,
    )


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
        ("입사~6개월 (즉시 기여)", ACCENT_BLUE),
        ("6개월~2년 (과제 확장)", NAVY),
        ("2년~ (Lab 비전 주도)", RGBColor(0x0A, 0x1A, 0x6E)),
    ]
    col_w = Inches(3.5)
    col_h = Inches(0.45)
    label_w = Inches(1.8)
    hdr_top = Inches(1.6)

    for i, (title, color) in enumerate(cols):
        cx = label_w + Inches(0.1) + Emu(i * (int(col_w) + int(Inches(0.15))))
        add_rect(slide, cx, hdr_top, col_w, col_h, fill_color=color)
        add_textbox(slide, cx, hdr_top, col_w, col_h,
                    title, font_size=10, font_color=WHITE, bold=True,
                    alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # 4 row labels
    row_labels = ["Eclipsa Audio", "AI 전환", "파트 시너지", "인증·표준"]
    row_h = Inches(1.0)
    row_start = hdr_top + col_h + Inches(0.1)

    grid = [
        # Eclipsa Audio
        [
            "청감 평가 프레임 구축\n(ISO 12913 + SATP 방법론)",
            "HRTF 개인화 알고리즘\n렌더러 품질 최적화",
            "차세대 Eclipsa Audio\n품질 표준 주도",
        ],
        # AI 전환
        [
            "A-JEPA 프로토타입\n음질 자동 판정 파이프라인",
            "생리신호 기반\n실시간 품질 모니터링",
            "AI 청감 평가 플랫폼\n자동화 완성",
        ],
        # 파트 시너지
        [
            "Display팀 협업\nA-V 통합 평가 설계",
            "Holographic Display ×\nSpatial Audio 시너지",
            "크로스모달 경험 설계\nLab 비전 제안",
        ],
        # 인증·표준
        [
            "Eclipsa Audio 인증 기준\n초안 작성",
            "국제 표준화 기여\n(IEC/ISO WG 참여)",
            "삼성 주도 표준\nde facto 확립",
        ],
    ]

    for r, (label, cells) in enumerate(zip(row_labels, grid)):
        ry = row_start + Emu(r * (int(row_h) + int(Inches(0.1))))
        # Row label
        add_rect(slide, Inches(0.1), ry, label_w, row_h, fill_color=LIGHT_BG)
        add_textbox(slide, Inches(0.1), ry, label_w, row_h,
                    label, font_size=10, font_color=NAVY, bold=True,
                    alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
        # Cells
        for c, cell_text in enumerate(cells):
            cx = label_w + Inches(0.1) + Emu(c * (int(col_w) + int(Inches(0.15))))
            add_rect(slide, cx, ry, col_w, row_h, fill_color=LIGHT_BG)
            add_textbox(slide, cx + Inches(0.1), ry + Inches(0.05),
                        col_w - Inches(0.2), row_h - Inches(0.1),
                        cell_text, font_size=9, font_color=DARK_TEXT,
                        alignment=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)

    # Bottom navy bar
    bar_top = Inches(6.4)
    add_rect(slide, Inches(0.6), bar_top, Inches(12.1), Inches(0.6), fill_color=NAVY)
    add_textbox(slide, Inches(0.8), bar_top + Inches(0.1), Inches(11.7), Inches(0.4),
                "신호처리가 만드는 기술 위에, 사용자가 경험하는 품질을 설계합니다",
                font_size=12, font_color=WHITE, bold=True,
                alignment=PP_ALIGN.CENTER)


# ============================================================================
# S18: Thank You
# ============================================================================
def build_s18(prs):
    slide = add_blank_slide(prs)
    set_slide_bg(slide, NAVY)

    # THANK YOU
    add_textbox(slide, Inches(0.8), Inches(2.0), Inches(11.5), Inches(1.0),
                "THANK YOU",
                font_size=36, font_color=WHITE, bold=True,
                alignment=PP_ALIGN.CENTER, font_name=FONT_EN)

    # Subtitle
    add_textbox(slide, Inches(0.8), Inches(3.2), Inches(11.5), Inches(0.5),
                "경청해 주셔서 감사합니다. 질문 부탁드립니다.",
                font_size=14, font_color=WARM_GRAY, bold=False,
                alignment=PP_ALIGN.CENTER)

    # Accent blue line
    add_accent_bar(slide, Inches(5.0), Inches(4.0), Inches(3.3),
                   Inches(0.04), color=ACCENT_BLUE)

    # Name
    add_textbox(slide, Inches(0.8), Inches(4.5), Inches(11.5), Inches(0.4),
                "조현인 (Hyun In Jo, Ph.D.)",
                font_size=13, font_color=WHITE, bold=True,
                alignment=PP_ALIGN.CENTER)

    # Contact
    add_textbox(slide, Inches(0.8), Inches(5.0), Inches(11.5), Inches(0.4),
                "best2012@naver.com | 010-6387-8402 | linkedin.com/in/hyunin-jo",
                font_size=10, font_color=WARM_GRAY, bold=False,
                alignment=PP_ALIGN.CENTER, font_name=FONT_EN)

    # Org path
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
