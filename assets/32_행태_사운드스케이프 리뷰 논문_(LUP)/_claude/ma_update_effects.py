# -*- coding: utf-8 -*-
"""
Paper32 — 2026-09-18 추가 전문평가분(DB 경로 BORDERLINE 재확보 등)의 메타분석 편입 판정·효과 계산

★ 왜 별도 파일인가
  8월 인용추적분은 효과를 `ma_v2.py` 본문에 하드코딩했다(CT0025·CT0126·CT0184). 그 값이 어디서
  왔는지는 주석으로만 남아 재현·검증이 어려웠다. 이번에는 **원문 수치(표 번호) → 변환 → 판정**을
  이 파일 하나에 두고, 산출은 `fulltext/ma/ma_update_effects.csv` 로 내보낸다.
  주분석에 들어가는 효과는 기존 입력표(`ma_correlation_input.csv`)에 행으로 들어가며,
  이 스크립트는 그 행이 여기서 계산한 값과 같은지 **대조만** 한다(입력표를 고쳐 쓰지 않는다).

판정 규칙(analysis_rules.md)
  §1 rho 는 r 근사로 취급(④ rho 제외 민감도 대상) · SE → SD = SE·√n (323 선례)
  §2 클러스터 부호 — MA1 g = (긍정·조용 조건 − 소음 조건)/SD (음수 = 소음에서 빠름)
                    MA3 g > 0 = 조용·자연음 조건에서 상호작용 많음
                    MA4 = 보고된 부호 그대로(8월 16편과 같은 처리. 부호 규칙 부재는 심사 M1 로 계류 중)
  §3 논문당 1효과 · 동급 다중 = 논문 내 평균(ρ = 0.5, Borenstein) — 사이트·실험이 서로 독립 표본이어도
     8월 선례(532 Exp1·Exp2)와 같이 ρ = 0.5 를 쓴다. ρ = 0 역분산 합성은 대안으로 함께 기록한다.
  §4 참가자 n 사용, 관측 n 뿐이면 민감도 전용 · n 미보고 → 제외
  §5 변환 경로가 규약 밖이면 제외(또는 변형분석)
  D5-4 "애매하면 넣지 않는다 — 넣었다면의 값은 민감도로 계산만"
  D5-5 표본 공유 논문은 같은 클러스터에 한 편만

출력: fulltext/ma/ma_update_effects.csv
"""
import sys, os, csv, math

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
MA = os.path.join(BASE, "fulltext", "ma")
sys.path.insert(0, BASE)
from ma_core import combine_within_study   # noqa: E402

CHINN = math.sqrt(3) / math.pi


def z_of(r):
    return math.atanh(r)


def site_composite(pairs, rho=0.5):
    """[(n, r)] → (z̄, var) — 규약 §3 논문 내 평균."""
    return combine_within_study([(z_of(r), 1 / (n - 3)) for n, r in pairs], rho=rho)


def site_iv(pairs):
    """[(n, r)] → 역분산(고정효과) 합성 — 독립 표본이면 이것이 표준(ρ = 0)."""
    w = [n - 3 for n, _ in pairs]
    z = sum(wi * z_of(r) for wi, (_, r) in zip(w, pairs)) / sum(w)
    return z, 1 / sum(w)


def or_2x2(a, b, c, d):
    """a,b = 조건1 (사건, 비사건) · c,d = 조건2 → Chinn d, var"""
    return math.log((a * d) / (b * c)) * CHINN, (1 / a + 1 / b + 1 / c + 1 / d) * CHINN ** 2


ROWS = []


def add(uid, cluster, decision, variant, label, source, values, n_info, method, y=None, v=None,
        is_r=False, reason=""):
    r = lo = hi = None
    if y is not None:
        se = math.sqrt(v)
        lo, hi = y - 1.96 * se, y + 1.96 * se
        if is_r:
            r, lo, hi = math.tanh(y), math.tanh(lo), math.tanh(hi)
    ROWS.append(dict(uid=uid, cluster=cluster, decision=decision, variant=variant, label=label,
                     source=source, values_verbatim=values, n_info=n_info, method=method,
                     y="" if y is None else f"{y:.10f}", v="" if v is None else f"{v:.10f}",
                     r="" if r is None else f"{r:.10f}",
                     lo95="" if lo is None else f"{lo:.4f}", hi95="" if hi is None else f"{hi:.4f}",
                     reason=reason))


# ═══════════════ MA4 소리–행태 상관 ═══════════════
# ── 80 Yu & Kang (2008) JASA 123:772 — Table IX "Moving activities" 열 × Table I 사이트별 응답자 수
#    이동 활동 = 에너지 소비 순 5범주(1 앉기·2 서기·3 걷기·4 아이와 놀기·5 운동), 조사원 관찰(원문 §II)
#    음량 평가 = −2 very quiet ~ +2 very noisy. Spearman 양측. 19개 사이트는 서로 다른 응답자(독립 표본).
#    검산: 사이트 n 으로 역산한 p 가 표의 p 와 모두 맞는다(예: 8번 r −.12, n 599 → p .003, 표 .00;
#    2번 r −.10, n 406 → p .044, 표 .04; 4번 r .06, n 848 → p .08, 표 .06).
SITES_80 = [  # (Table I 응답자 수, Table IX rho)
    (418, -0.01), (406, -0.10), (655, 0.05), (848, 0.06), (777, 0.03), (1037, -0.03),
    (574, 0.02), (599, -0.12), (888, -0.04), (1041, 0.00), (459, 0.05), (489, 0.06),
    (499, -0.03), (510, 0.02), (307, -0.04), (304, -0.06), (62, -0.05), (79, -0.05), (79, 0.00)]
V80 = ("Table IX Moving activities rho by site 1–19: " +
       "; ".join(f"{i}: {r:+.2f}" for i, (_, r) in enumerate(SITES_80, 1)))
N80 = "Table I interviewees per site " + ", ".join(str(n) for n, _ in SITES_80) + \
      f" (합 {sum(n for n, _ in SITES_80):,}) — 한 응답자 1관측, 항목 결측 미보고"
z80, v80 = site_composite(SITES_80)
add("80", "MA4", "main", "primary", "Yu & Kang 2008 (이동 활동 강도 ↔ 음량 평가, 19개 사이트)",
    "Table IX; Table I; §II", V80, N80,
    "사이트별 rho → Fisher z(v = 1/(n−3)) → 규약 §3 논문 내 평균(ρ = 0.5, 532 선례). 부호 = 보고 그대로",
    z80, v80, True,
    "관찰 행태 × 현장 음 지각, 참가자 단위, 규약 내 변환 — MA4 정의(음향·지각 지표 × 행태, 방향 무관) 충족. "
    "표본 공유(104·123)는 D5-5 로 80 한 편만(관찰 기록 명시 = §3 우선순위 ①)")
z, v = site_iv(SITES_80)
add("80", "MA4", "alt", "iv_rho0", "80 대안: 사이트 역분산 합성(ρ = 0)", "Table IX; Table I", "", "",
    "독립 표본 표준 합성 — Σ(n−3)z/Σ(n−3), v = 1/Σ(n−3)", z, v, True,
    "규약 문언(ρ = 0.5)보다 정밀. 랜덤효과 가중치는 τ² 가 지배해 결과 차이 미미")
add("80", "MA4", "alt", "sign_flip", "80 대안: 부호 반전(조용 평가 ↔ 활동 강도)", "Table IX", "", "",
    "주 합성값의 부호만 뒤집음", -z80, v80, True,
    "MA4 에 부호 규칙이 없으므로(심사 M1) 방향 가정이 결과를 바꾸는지 확인")

# ── 104 Yu & Kang (2009) JASA 126:1163 — 같은 19개 사이트 설문(Table I 응답자 수 동일, B4 값이 80 과 일치)
#    Table III. txt 추출본에 음수 부호가 없어 추출 단계에서 원 PDF 글리프로 복원(newft_final/es.csv 메모).
B5 = [0.05, 0.08, 0.00, -0.10, -0.11, 0.10, 0.15, 0.01, -0.07, 0.00, -0.06, -0.00, -0.00, -0.04,
      -0.10, -0.01, 0.13, 0.05, -0.05]
B7 = [0.01, 0.03, 0.07, 0.12, -0.03, -0.02, -0.04, 0.10, 0.00, -0.09, 0.03, -0.05, 0.06, 0.10,
      0.09, -0.10, 0.00, 0.11, -0.04]
NS = [n for n, _ in SITES_80]
for tag, vals, lab in (("B5", B5, "방문 빈도 ↔ 음량 평가"), ("B7", B7, "동행 여부 ↔ 음량 평가")):
    z, v = site_composite(list(zip(NS, vals)))
    add("104", "MA4", "overlap", tag, f"Yu & Kang 2009 ({lab}, 19개 사이트)", "Table III; Table I",
        f"{tag} by site: " + "; ".join(f"{x:+.2f}" for x in vals), "80 과 같은 응답자",
        "사이트별 rho → 논문 내 평균(ρ = 0.5) — 참고용", z, v, True,
        "80 과 동일 표본(D5-5 한 편만). B4(이동 상태)는 80 의 Table IX 값과 같다. "
        "80 을 택한 이유 = 행태를 조사원이 관찰했다고 명시(§3 ①). 부호는 PDF 글리프 복원값")

# ── CT0171 Yang, Cao & Meng (2021 온라인; STOTEN 802:149869, 2022) — Table 1
#    군중밀도(드론 사진, 30 m × 30 m) × RSI(인간음/자연음·교통음 지각 비율). Spearman.
#    ⚠️ 군중밀도는 "측정점별 평균"(§3.1.1) — RSIn 은 7개 측정점·42장, RSIt 는 2개 측정점·12장.
#       N(67·56)은 응답자 수라 행태 변수의 독립 단위(측정점)보다 크다. CT0126 은 같은 문제에서
#       분석단위를 측정점(N = 29)으로 잡았다(D5-5) — 같은 원칙이면 RSIt(측정점 2)는 상관 자체가 성립하지 않는다.
zA, vA = z_of(0.302), 1 / (67 - 3)
zB, vB = z_of(0.553), 1 / (56 - 3)
z, v = combine_within_study([(zA, vA), (zB, vB)])
add("CT0171", "MA4", "sensitivity", "primary", "Yang et al. 2021 (군중밀도 ↔ 인간음 지각 비율)",
    "Table 1; §3.1.1", "RSIn rho = 0.302* (N = 67); RSIt rho = 0.553** (N = 56)",
    "N = 응답자(Table 1 주). 군중밀도는 측정점 7개(42장)·2개(12장)의 평균",
    "rho → z, n 은 보고 N · 두 하위표본 논문 내 평균(ρ = 0.5)", z, v, True,
    "애매 → 주분석 제외. 행태 변수(군중밀도)가 측정점 단위(7·2곳)인데 N 은 응답자라 정밀도가 과대 — "
    "CT0126(측정점 N = 29 채택) 원칙과 충돌. 규약 §4 관측 n 문제와 같은 구조")
z, v = site_iv([(67, 0.302), (56, 0.553)])
add("CT0171", "MA4", "alt", "iv_rho0", "CT0171 대안: 역분산 합성(ρ = 0)", "Table 1", "", "",
    "두 하위표본 역분산", z, v, True, "")

# ── 1305 (2026) German J Exerc Sport Res — Table 4 "Natural sounds" 열, n = 408 완결 사례
R1305 = [("precipitation", 0.151), ("cold", 0.170), ("strong wind", 0.076), ("heat", 0.051),
         ("sun radiation", 0.034)]
z, v = combine_within_study([(z_of(r), 1 / (408 - 3)) for _, r in R1305])
add("1305", "MA4", "sensitivity", "primary", "Reuß & Huth 2026 (자연음 중요도 ↔ 악천후 그린운동)",
    "Table 4; §Data analysis", "; ".join(f"{k} {r:.3f}" for k, r in R1305),
    "408명(772 응답 중 결측 제외), 상관별 n 미보고",
    "rho → z(n = 408) · 같은 표본 5효과 논문 내 평균(ρ = 0.5)", z, v, True,
    "애매 → 주분석 제외. 소리 쪽 변수가 장소의 음환경 지각이 아니라 '자연음이 나에게 중요한가'(태도·선호) "
    "5점 문항 — MA4 의 음향·지각 지표 정의 밖. 판정 신뢰도도 low")

# ── 713 Ma et al. (2021) Applied Acoustics 171:107570 — Table 7 단계적 회귀
#    y = 좋은 사운드스케이프에 대한 전반적 선호(1–5), 방문빈도 B = 0.17 (SE 0.07), β = 0.18, p = .017,
#    n = 150, 최종 모형 예측변수 3개(쾌적성 지각·자연음 선호 점수·방문빈도)
t713 = 0.17 / 0.07
df713 = 150 - 3 - 1
rp = t713 / math.sqrt(t713 ** 2 + df713)
add("713", "MA4", "sensitivity", "primary", "Ma et al. 2021 (방문빈도 → 좋은 사운드스케이프 선호)",
    "Table 7", "Visit frequency B 0.17 (SEB 0.07), 95% CI [0.03, 0.32], b 0.18, p 0.017; F(3,146)",
    "150명(5개 지점 × 30)", "규약 §1 표준화 β+SE → 부분상관 근사: t = B/SE, r_p = t/√(t²+df), "
    "v = 1/(n − 3 − 2공변량)", z_of(rp), 1 / (150 - 3 - 2), True,
    "애매 → 주분석 제외. 결과가 장소 음환경 지각이 아니라 '좋은 사운드스케이프 선호' 평정이고, "
    "쾌적성 지각을 통제한 단계적 회귀의 부분효과(규약 §1 '주의 라벨' 경로). 추출 단계 판정도 계산 불가")

# ── 계산 불가
add("499", "MA4", "not_computable", "", "Song et al. 2018 (방문빈도 ↔ 음 요소 선호)", "Table 3",
    "V1 0.263* … V12 −0.071 (12개 음 요소)", "1,260명 중 상관에 쓰인 n 미보고",
    "", reason="n 미보고(규약 §4). n = 1,260 이면 r = .110 도 p < .001 이어야 하는데 유의 표기가 없어 "
    "분석 단위 불명. 소리 쪽도 선호(태도) 변수")
add("1218", "MA4", "not_computable", "", "Huang et al. 2026 (음원 유형 ↔ SOPARC 행태)",
    "§3.2.2 Fig. 12; Table 3–4", "상관계수는 그림에만. 회귀 RS β 0.627 (SE 0.098), 교호작용 포함 모형",
    "27개 모니터링 지점 × 시간 스캔, 모형 n 미보고",
    "", reason="상관계수 미보고(그림만). 회귀는 교호작용 항이 있는 모형의 조건부 계수이고 n 미보고 — "
    "adj R² 에서 역산하면 n ≈ 53 이나 원문에 없는 값을 만드는 것이라 쓰지 않는다(D7-3 1018 교훈)")
add("CT0223", "MA4", "not_computable", "", "Kuldna et al. 2019 (체류 예정시간 → 자연음 청취 만족)",
    "Table 5", "1–2 h 0.11, 2–5 h 0.42, >5 h 1.35** (ordered logit 계수)", "n = 528",
    "", reason="SE·CI 미보고, 다범주 순서 로짓 계수 — 규약 §1 경로 없음. 결과도 만족(지각)")

# ═══════════════ MA1 보행속도 ═══════════════
# ── 709 Levenhagen et al. (2021) People and Nature — §3.3
#    '정숙' 안내판 있음/없음 주간 교대(10주). 보행속도 sign absent 1.03 ± 0.02 m/s (n = 958),
#    sign present 1.01 ± 0.02 m/s (n = 974). 같은 논문의 다른 수치가 "mean ± SE" 로 명시돼 있어 ± 는 SE.
#    Kruskal–Wallis χ² = 3.2, df = 1, p = 0.08. 이상치 1건 제외. 안내판은 L50 을 1.19 dB(A) 낮췄다.
m_pos, m_neg, se_pos, se_neg, n_pos, n_neg = 1.01, 1.03, 0.02, 0.02, 974, 958
sd_pos, sd_neg = se_pos * math.sqrt(n_pos), se_neg * math.sqrt(n_neg)
df = n_pos + n_neg - 2
sp = math.sqrt(((n_pos - 1) * sd_pos ** 2 + (n_neg - 1) * sd_neg ** 2) / df)
J = 1 - 3 / (4 * df - 1)
g709 = J * (m_pos - m_neg) / sp
v709 = (n_pos + n_neg) / (n_pos * n_neg) + g709 ** 2 / (2 * (n_pos + n_neg))
add("709", "MA1", "sensitivity", "primary", "Levenhagen et al. 2021 (정숙 안내판 있음 vs 없음)",
    "§3.3 Results; §3.1", "sign absent 1.03 ± 0.02 m/s (958) · present 1.01 ± 0.02 m/s (974); "
    "KW χ² = 3.2, df = 1, p = 0.08", "보행속도 관측 1,932건(개인 식별 후 계시, 같은 방문객 중복 여부 미보고)",
    "SE → SD = SE·√n (323 선례) → Hedges g, 부호 = (조용 조건 − 소음 조건)", g709, v709, False,
    "애매 → 주분석 제외. 대비가 음환경 조건이 아니라 '정숙' 안내판(행동 규범 개입)이라 속도 차이를 소리로 "
    "귀속할 수 없다 — CT0414(대비 불일치) 선례. 평균이 소수 둘째 자리 반올림이라 g 의 불확실성이 크다")
r_kw = math.sqrt(3.2 / (n_pos + n_neg - 1))
d_kw = -2 * r_kw / math.sqrt(1 - r_kw ** 2)   # 방향: 안내판 조건이 느림
add("709", "MA1", "alt", "kruskal", "709 대안: KW χ²(1) → r → d", "§3.3", "", "",
    "r = √(χ²/N), d = 2r/√(1−r²), 부호는 평균 방향", d_kw,
    (n_pos + n_neg) / (n_pos * n_neg) + d_kw ** 2 / (2 * (n_pos + n_neg)), False,
    "순위 검정 경로(규약 §5 는 U+n 완비 시만 허용 — 참고값)")
add("OAS0005", "MA1", "not_computable", "", "Zeng et al. 2018 (단일열 보행 실험, 배경음악 120 BPM)",
    "Fundamental diagram §", "밀도 구간별 속도 감소율 2.97%·1.40%·20.24%·37.89%", "40명, 실험실",
    "", reason="SENS_ONLY(P1 실험실 재현). 속도는 % 변화만 있고 SD 없음 — 계산 불가. 대비도 음악(481 선례상 주분석 밖)")

# ═══════════════ MA3 사회적 상호작용 — 등록 민감도 (3) 행태 의도 연구 포함 ═══════════════
# ── 917 Su, Ma & Wang (2023) Applied Acoustics 207:109350 — Table 3(분할표)
#    아동 38명, 9개 음 조건 × 2개 시각 장면 반복측정(조건당 관측 65–66건). 결과 = 행동 기대(의도).
#    MA3 의 기존 대비(자연음 vs 교통·공사 소음, 931·1069)와 같은 짝: Nature vs Motorized transport,
#    Nature vs Electro-mechanical(공사·환기) — 규약 §3 동급 2효과 → ρ = 0.5 평균.
#    Interaction vs (Nonparticipation + Avoiding) 이분화. g > 0 = 자연음에서 상호작용 기대 많음.
NAT = (41, 65 - 41)
MOT = (23, 66 - 23)
ELE = (18, 66 - 18)
NOS = (24, 66 - 24)
dM, vM = or_2x2(*NAT, *MOT)
dE, vE = or_2x2(*NAT, *ELE)
d917, v917obs = combine_within_study([(dM, vM), (dE, vE)])
DEFF = 66 / 38   # 조건당 관측 66 / 참가자 38 — 규약 §4 참가자 n 으로 되돌리는 분산 팽창
add("917", "MA3", "sensitivity", "primary", "Su et al. 2023 (자연음 vs 교통·기계음, 상호작용 기대)",
    "Table 3", "Nature 41/65 · Motorized 23/66 · Electro-mechanical 18/66 · No sound 24/66 (Interaction/total)",
    "아동 38명 반복측정, 조건당 관측 65–66", "2×2 OR → Chinn d(자연 vs 교통, 자연 vs 기계) → ρ = 0.5 평균 → "
    f"분산 × {DEFF:.3f}(관측 66 → 참가자 38, 규약 §4)", d917, v917obs * DEFF, False,
    "SENS_ONLY(R2 행동 의도). 등록 민감도 (3) '행태 의도 연구 포함' 대상. 실험실 VR 재현이기도 함")
add("917", "MA3", "alt", "obs_n", "917 대안: 관측 n 그대로", "Table 3", "", "", "위와 같되 분산 보정 없음",
    d917, v917obs, False, "관측 단위를 독립으로 취급 — 규약 §4 상 참고값")
dQ1, vQ1 = or_2x2(*NOS, *MOT)
dQ2, vQ2 = or_2x2(*NOS, *ELE)
dq, vq = combine_within_study([(dQ1, vQ1), (dQ2, vQ2)])
add("917", "MA3", "alt", "nosound_vs_noise", "917 대안: 무음 vs 교통·기계음(조용 대비 프레임)", "Table 3", "", "",
    "무음 대조를 '조용' 조건으로 본 경우(Moser·Mathews 형 대비)", dq, vq * DEFF, False,
    "실험실의 무음은 현장의 조용한 조건과 다르다(아동이 무음 조건을 가장 많이 회피) — 참고값")


def main():
    out = os.path.join(MA, "ma_update_effects.csv")
    cols = ["uid", "cluster", "decision", "variant", "label", "source", "values_verbatim", "n_info",
            "method", "y", "v", "r", "lo95", "hi95", "reason"]
    with open(out, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=cols)
        w.writeheader(); w.writerows(ROWS)

    # ── 주분석 편입분이 입력표에 같은 값으로 들어가 있는지 대조 ─────────────────
    main_rows = [r for r in ROWS if r["decision"] == "main"]
    corr = {r["no"]: r for r in csv.DictReader(open(os.path.join(MA, "ma_correlation_input.csv"),
                                                   encoding="utf-8-sig"))}
    bad = []
    for m in main_rows:
        c = corr.get(m["uid"])
        if not c or (c.get("status") or "include") != "include":
            bad.append(f"{m['uid']} 가 ma_correlation_input.csv 에 include 로 없다"); continue
        if abs(math.atanh(float(c["r"])) - float(m["y"])) > 1e-7 or abs(float(c["v_override"]) - float(m["v"])) > 1e-9:
            bad.append(f"{m['uid']} 입력표 값이 계산값과 다르다 (r {c['r']} vs {m['r']}, v {c['v_override']} vs {m['v']})")
    for r in ROWS:
        yv = f"{float(r['y']):+.3f}" if r["y"] else "—"
        rv = f" (r {float(r['r']):+.3f})" if r["r"] else ""
        print(f"{r['uid']:8s} {r['cluster']} {r['decision']:14s} {r['variant']:16s} y={yv}{rv} "
              f"v={r['v'][:8] if r['v'] else '—'}")
    print(f"[저장] {os.path.relpath(out, BASE)} ({len(ROWS)}행)")
    if bad:
        print("⚠️ " + " · ".join(bad))
        sys.exit(1)
    print("[대조] 주분석 편입분 = 입력표 ✅")


if __name__ == "__main__":
    main()
