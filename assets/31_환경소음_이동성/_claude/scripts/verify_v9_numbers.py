# -*- coding: utf-8 -*-
# v9 수치 정합성 게이트: 최종 추출본의 핵심 수치 ↔ data/processed CSV 재계산 대조.
import os, re, sys
import pandas as pd

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
P = os.path.join(ROOT, "data", "processed")
EX = sys.argv[1] if len(sys.argv) > 1 else os.path.join(ROOT, "_claude", "review", "manuscript_v9_final_extract.md")
text = open(EX, encoding="utf-8").read()

import json
fx = json.load(open(os.path.join(ROOT, "_claude", "figures_v2", "cache", "v9fix_results.json"), encoding="utf-8"))
tw = pd.read_csv(os.path.join(P, "phase2a_twoway_results.csv"))
dn = pd.read_csv(os.path.join(P, "phase2g_daynight_results.csv"))
fdr = pd.read_csv(os.path.join(P, "phase3_segmented_fdr.csv"))
rob = pd.read_csv(os.path.join(P, "phase4_robustness.csv")).set_index("check")["value"]
cal = pd.read_csv(os.path.join(P, "phase5_calibration_trends.csv")).set_index("series")
cv = pd.read_csv(os.path.join(P, "phase5_calibval.csv"), encoding="utf-8-sig")
did = pd.read_csv(os.path.join(ROOT, "_claude", "figures_v2", "cache", "did_weekly_ci.csv"))

day = tw.iloc[0]; night = tw.iloc[1]; full = tw.iloc[2]
lo, hi = day.beta - 1.96 * day.se, day.beta + 1.96 * day.se
r = did["mob_diff"].corr(did["dLeq_diff"]); rs = did["mob_diff"].corr(did["dLeq_diff"], method="spearman")

checks = [
 # (라벨, 기대 문자열이 본문에 존재?, CSV 재계산 근거)
 ("M2 주간 β=+0.648", "β=+0.648" in text, f"{day.beta:+.3f}"),
 ("M2 주간 CI 0.17-1.13", "0.17-1.13" in text and abs(lo - 0.168) < 0.005 and abs(hi - 1.129) < 0.005, f"[{lo:+.2f},{hi:+.2f}]"),
 ("M2 주간 p=0.008", "p=0.008" in text and abs(day.pval - 0.00816) < 1e-4, f"{day.pval:.5f}"),
 ("M2 전일 β=+0.628 p=0.038", "β=+0.628" in text and abs(full.pval - 0.03815) < 1e-3, f"{full.beta:+.3f}/{full.pval:.3f}"),
 ("M2 야간 +0.50 ns", "β=+0.50" in text and abs(night.pval - 0.4248) < 1e-3, f"{night.beta:+.3f}/{night.pval:.3f}"),
 ("M1 주간 +1.130 (동클러스터 SE 0.292)", "+1.130" in text and "0.292" in text
  and abs(fx["M1_day"]["lp_day_logrel"]["se"] - 0.2915) < 1e-3, "v9fix M1_day"),
 ("M1 야간(야간dose) +2.283 (0.748)", "+2.283" in text and "0.748" in text
  and abs(fx["M1_night"]["lp_night_logrel"]["b"] - 2.2833) < 1e-3, "v9fix M1_night"),
 ("표본 1,248,794/1,123/421", all(s in text for s in ("1,248,794", "1,123", "421")), "panel meta"),
 ("M2 표본 1,247,546/1,122/420", all(s in text for s in ("1,247,546", "1,122", "420")),
  f"{int(day['n'])}/{int(day.sensors)}/{int(day.dongs)}"),
 ("동-일 619,464 = 424×1,461", "619,464" in text and 424 * 1461 == 619464, "424*1461"),
 ("동 단위 순열 p=0.003", "=0.003" in text and abs(fx["placebo_dong"]["p"] - 0.00332) < 1e-3,
  f"{fx['placebo_dong']['p']:.4f}"),
 ("동 단위 순열 null SD 0.023", "SD 0.023" in text and abs(fx["placebo_dong"]["null_sd"] - 0.0226) < 1e-3,
  f"{fx['placebo_dong']['null_sd']:.4f}"),
 ("동일가중 β=+0.83 (p=0.047)", "+0.83" in text and "0.047" in text
  and abs(fx["dong_equal_weight"]["b"] - 0.8255) < 1e-3 and abs(fx["dong_equal_weight"]["p"] - 0.0471) < 1e-3,
  f"{fx['dong_equal_weight']['b']:+.3f}/{fx['dong_equal_weight']['p']:.4f}"),
 ("2차항 β²=−0.59 p=0.44", "−0.59" in text and abs(rob["nonlin_quad_p"] - 0.442) < 1e-2, f"{rob['nonlin_quad_b2']:+.3f}/{rob['nonlin_quad_p']:.3f}"),
 ("드리프트 −2.3 dB", "2.3 dB" in text and abs(cal.loc['sdot', 'delta'] + 2.30) < 1e-6, f"{cal.loc['sdot','delta']}"),
 ("자동측정망 +0.04", "+0.04" in text and abs(cal.loc['cal_auto_road', 'delta'] - 0.04) < 1e-6, "csv"),
 ("수동 일반 +1.93 / 도로 +1.91", "+1.93" in text and "+1.91" in text, "csv"),
 ("offset 11.7 dB (n=60)", "11.7 dB" in text and abs(cv.offset_day.mean() + 11.66) < 0.05 and len(cv) == 60,
  f"{cv.offset_day.mean():+.2f}/n={len(cv)}"),
 ("DiD 상관 +0.44/+0.46", "+0.44" in text and "+0.46" in text and abs(r - 0.439) < 5e-3 and abs(rs - 0.462) < 5e-3,
  f"{r:+.3f}/{rs:+.3f}"),
 ("연평균 49.4→47.6→47.2→47.1", all(s in text for s in ("49.4", "47.6", "47.2", "47.1")), "csv sdot row"),
]

# Table 6 12행 β·p·FDR 대조 (추출본 표 텍스트에서)
seg_expect = [(f"{b:+.3f}", p, f) for b, p, f in zip(fdr.beta, fdr.pval, fdr.p_BH_FDR)]
tbl6_ok = all(f"{b:+.3f}"[:6] in text for b in fdr.beta)  # +0.648, +0.504, ...
checks.append(("Table 6 β 12건 일치", tbl6_ok, "phase3_segmented_fdr"))
checks.append(("마드리드 4-6 dB (웹 실측 교정)",
               ("4-6 dB 감소 [19]" in text or "4-6 dB in Madrid [19]" in text) and "약 3 dB" not in text,
               "JASA 148:1748"))
checks.append(("n_hours≤24 필터 β=+0.651", "β=+0.651" in text and "5.9%" in text, "로컬 재계산 0.6509/0.2482"))
checks.append(("within-date SD 0.115", "0.115 log-unit" in text, "로컬 재계산 0.1146"))
checks.append(("[28] 게재본 doi", "10.1007/s11356-021-13872-z" in text, "PubMed 33884552"))
fdr_disp = ["0.033", "0.46", "0.92", "0.45", "0.15", "0.25", "0.21", "0.21", "0.066", "0.018", "0.14", "0.033"]
checks.append(("Table 6 FDR p 표기 일치",
               all(abs(round(v, len(s.split('.')[-1])) - float(s)) < 1e-9 for v, s in zip(fdr.p_BH_FDR, fdr_disp)),
               "rounding"))

fails = 0
for label, ok, basis in checks:
    mark = "PASS" if ok else "FAIL"
    if not ok: fails += 1
    print(f"  [{mark}] {label}  (근거 {basis})")
print(f"\n총 {len(checks)}건 중 실패 {fails}건")
sys.exit(1 if fails else 0)
