# -*- coding: utf-8 -*-
# 생활인구 partial_*.csv(seq별 동×일 집약) 병합 -> livingpop_dong_daily_2020-2023.csv (원시 코드 보존).
# 크로스워크(강북 코드 relabel·상일 병합)는 패널 빌더(build_analysis_panel)에서 dose_key로 적용.
import csv, os, glob
from collections import defaultdict

TMP = r"C:\Users\wh850\AppData\Local\Temp\livingpop"
ROOT = r"T:\00_BACKUP\01_한양대학교\00_연구실\00_프로젝트\★_논문화 프로젝트\31_환경소음_이동성"
OUT = os.path.join(ROOT, r"data\processed\livingpop_dong_daily_2020-2023.csv")

parts = sorted(glob.glob(os.path.join(TMP, "partial_*.csv")))
print(f"partial 파일: {len(parts)}")
seen = {}  # (adm_cd,date) -> row  (중복 검출)
dups = 0
for p in parts:
    with open(p, encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            k = (r["adm_cd"], r["date"])
            if k in seen:
                dups += 1
            else:
                seen[k] = (r["lp_mean"], r["lp_day"], r["lp_night"], r["n_hours"])

rows = sorted(seen.items(), key=lambda kv: (kv[0][0], kv[0][1]))
with open(OUT, "w", newline="", encoding="utf-8-sig") as f:
    w = csv.writer(f)
    w.writerow(["adm_cd", "date", "lp_mean", "lp_day", "lp_night", "n_hours"])
    for (adm, dt), (m, d, nt, nh) in rows:
        w.writerow([adm, dt, m, d, nt, nh])

dongs = sorted(set(k[0] for k in seen))
days = sorted(set(k[1] for k in seen))
print(f"행: {len(rows)} | 동: {len(dongs)} | 일: {len(days)} ({days[0]}~{days[-1]}) | 중복: {dups}")
print(f"동×일 기대치(완전): {len(dongs)*len(days)} vs 실제 {len(rows)}  결측 {len(dongs)*len(days)-len(rows)}")
print(f"-> {OUT}")
