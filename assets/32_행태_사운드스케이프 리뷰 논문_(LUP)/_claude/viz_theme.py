# -*- coding: utf-8 -*-
"""
Paper32 — 그림 공통 테마 v2 (2026-08-16 전면 개편)

★ 왜 바꿨나 (사용자 지적 3건)
  ① 가독성: 그림을 7.5~8.6 in 로 그려 놓고 docx 에는 5.0~5.9 in 로 축소 삽입
     → 인쇄 실효 글자가 4~6 pt 까지 떨어졌다. **작화 폭 = 삽입 폭(W_FULL)** 로 통일해
     그린 글자 크기가 곧 인쇄 크기가 되게 한다.
  ② 색: 구 팔레트(#2a78d6·#eb6834)는 채도가 높아 인쇄 지면에서 요란했다.
     저채도 슬레이트 블루·테라코타·웜 그레이로 교체. 의미 배정은 유지
     (blue = forward/긍정 · terracotta = reverse/부정 · gray = 중립/판단불가).
  ③ 비율: 각 그림 스크립트에서 내용에 맞게 재설계(Fig7 정사각 2단 → 가로 병렬 등).

검증 (validate_palette 재현 · OKLab ΔE×100 · Viénot CVD 시뮬 · 2026-08-16):
  blue #34608D vs terracotta #B4593F : normal 22.3 / deuter 33.3 / protan 28.6  PASS
  blue vs gray #A9A29A               : normal 25.8 / deuter 32.2 / protan 30.1  PASS
  terracotta vs gray                 : normal 18.5 / deuter 10.6 / protan 11.2  PASS
  순서형 램프 #93B2D1/#4E7AA6/#24476B: 인접쌍 normal 18.7·17.7 / CVD 전부 ≥16  PASS
  gray 대비 2.52:1 < 3 → relief rule: 회색 세그먼트에는 항상 수치 라벨을 직접 인쇄한다.
"""
import matplotlib as mpl
from matplotlib.colors import LinearSegmentedColormap, to_rgb

# ── 카테고리(고정 순서 — 순환 금지) ────────────────────────────────
BLUE = "#34608D"     # slate blue   — forward · 긍정 · 1차 계열
TERRA = "#B4593F"    # terracotta   — reverse · 부정
NEUT = "#A9A29A"     # warm gray    — 양방향 · 판단불가 · 중립
SERIES = [BLUE, TERRA, NEUT]
ORANGE, MUTED = TERRA, NEUT          # 구 코드 호환 별칭

# 순서형 3단 (측정세대 G1<G2<G3 · 품질 low<moderate<high)
ORD3 = ["#93B2D1", "#4E7AA6", "#24476B"]
DEEP = "#24476B"     # 강조(풀링 다이아몬드·합계 등)

# 도면 상자 틴트(프레임워크·PRISMA)
TINT_BLUE = "#E9EFF6"
TINT_TERRA = "#F6EAE4"
TINT_NEUT = "#F5F4F1"

# ── 잉크·크롬 ──────────────────────────────────────────────────────
INK = "#1A1A1A"      # primary
INK2 = "#55524E"     # secondary
AXIS = "#8C8880"     # muted (축·주석)
GRID = "#E7E5E0"     # hairline
BASE = "#C6C2BB"     # baseline·상자 테두리(약)
SURF = "#FFFFFF"     # 인쇄 대비 백색 표면

# ── 단일 hue 순차 램프 (히트맵 — 0 은 표면으로 후퇴) ────────────────
SEQ_STEPS = ["#F2F6FA", "#DCE7F1", "#C2D5E6", "#A3BFD8", "#7FA4C4",
             "#5B84AC", "#3E6890", "#2C5379", "#1C3A57"]
SEQ = LinearSegmentedColormap.from_list("p32_slate", SEQ_STEPS)

FRAME_POS = DEEP
FRAME_NEG = TERRA

# ── 판형: 작화 폭 = docx 삽입 폭 (A4 여백 제외 6.09 in) ─────────────
W_FULL = 6.05


def ink_on(color):
    """배경색 위 글자색 — 눈대중이 아니라 휘도로 결정한다."""
    r, g, b = to_rgb(color)
    lum = 0.2126 * r + 0.7152 * g + 0.0722 * b   # 근사(감마 생략, 임계용으로 충분)
    return SURF if lum < 0.52 else INK


def apply():
    mpl.rcParams.update({
        "font.family": ["Arial", "Helvetica", "DejaVu Sans"],
        "font.size": 8,
        "text.color": INK,
        "axes.edgecolor": BASE,
        "axes.labelcolor": INK2,
        "axes.linewidth": 0.8,
        "axes.titlecolor": INK,
        "xtick.color": AXIS, "ytick.color": AXIS,
        "xtick.labelcolor": INK2, "ytick.labelcolor": INK2,
        "xtick.major.width": 0.7, "ytick.major.width": 0.7,
        "xtick.labelsize": 7.5, "ytick.labelsize": 7.5,
        "grid.color": GRID, "grid.linewidth": 0.6,
        "figure.facecolor": SURF, "axes.facecolor": SURF,
        "savefig.facecolor": SURF,
        "savefig.dpi": 300, "savefig.bbox": "tight",
        "legend.frameon": False,
    })
