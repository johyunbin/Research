# -*- coding: utf-8 -*-
# 생활인구 행정동(OA-14991) zip 1개 → 동×일 집약 partial CSV (append).
# 32컬럼: [0]기준일ID [1]시간대구분(00-23) [2]행정동코드(8자리) [3]총생활인구수 [4:]성연령별
# 총생활인구수=시간별 체류(stock) → 일 대표=시간 평균(합 아님). 주간 06-21 / 야간 22-05 분리.
# 사용: python livingpop_to_daily.py <zip_path> <partial_out.csv>
import sys, os, zipfile, io, csv
from collections import defaultdict

zip_path, out = sys.argv[1], sys.argv[2]
DATE_I, HOUR_I, DONG_I, POP_I = 0, 1, 2, 3

z = zipfile.ZipFile(zip_path)
csvs = [n for n in z.namelist() if n.lower().endswith(".csv")]
acc = defaultdict(lambda: [0.0, 0, 0.0, 0, 0.0, 0])  # (date,dong)->[s24,n24,sday,nday,snt,nnt]

for n in csvs:
    with z.open(n) as fp:
        rdr = csv.reader(io.TextIOWrapper(fp, encoding="utf-8-sig", newline=""))
        next(rdr, None)  # header
        for r in rdr:
            if len(r) <= POP_I:
                continue
            d = r[DATE_I].strip()
            if len(d) == 8 and d.isdigit():
                d = f"{d[:4]}-{d[4:6]}-{d[6:8]}"
            dong = r[DONG_I].strip()
            pv = r[POP_I].strip()
            if not pv:
                continue
            try:
                p = float(pv)
            except ValueError:
                continue
            try:
                hh = int(r[HOUR_I].strip())
            except ValueError:
                hh = None
            a = acc[(d, dong)]
            a[0] += p; a[1] += 1
            if hh is not None:
                if 6 <= hh <= 21:
                    a[2] += p; a[3] += 1
                else:
                    a[4] += p; a[5] += 1

def m(s, n):
    return round(s / n, 1) if n else ""

newfile = not os.path.exists(out)
with open(out, "a", newline="", encoding="utf-8-sig") as f:
    w = csv.writer(f)
    if newfile:
        w.writerow(["adm_cd", "date", "lp_mean", "lp_day", "lp_night", "n_hours"])
    for (d, dong), a in sorted(acc.items()):
        w.writerow([dong, d, m(a[0], a[1]), m(a[2], a[3]), m(a[4], a[5]), a[1]])

dongs = len(set(k[1] for k in acc))
days = len(set(k[0] for k in acc))
print(f"{os.path.basename(zip_path)}: files={len(csvs)} dongs={dongs} days={days} dong-days={len(acc)} -> appended {out}")
