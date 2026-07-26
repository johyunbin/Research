# -*- coding: utf-8 -*-
# 2023 스키마 검증 + 일패널 연도간 분포 정합성
import zipfile, io, csv, os
from collections import defaultdict
build = r"C:\Users\wh850\AppData\Local\Temp\sdot_build"
def decode(b):
    for e in ("cp949","euc-kr","utf-8-sig","utf-8"):
        try: return b.decode(e)
        except: pass
    return b.decode("utf-8","replace")

# --- 2023 헤더: '소음' 포함 컬럼 전부 + idx24 값 샘플 ---
z=zipfile.ZipFile(os.path.join(build,"sdot_2023.zip"))
csvs=[n for n in z.namelist() if n.lower().endswith(".csv")]
rdr=csv.reader(io.StringIO(decode(z.read(csvs[0])))); h=next(rdr)
print(f"2023 header cols={len(h)}")
print(" '소음' 포함 컬럼:")
for i,c in enumerate(h):
    if "소음" in (c or ""): print(f"    idx{i}: {c!r}")
print(" idx1/idx24/idx57:", repr(h[1]), repr(h[24]), repr(h[57]))
print(" 샘플 row 값 idx24(소음):", end=" ")
for k,r in enumerate(rdr):
    if k>=5: break
    print(r[24] if len(r)>24 else "?", end="  ")
print()

# --- 일패널 연도별 분포 ---
panel=os.path.join(build,"sdot_daily_panel.csv")
by=defaultdict(list); gus=defaultdict(set); sensors=defaultdict(set)
for row in csv.DictReader(open(panel,encoding="utf-8-sig")):
    y=row["date"][:4]
    try: by[y].append(float(row["Leq24"]))
    except: pass
    if row["gu"]: gus[y].add(row["gu"])
    sensors[y].add(row["serial"])
print("\nyear | sensor-days | Leq24 p05/median/p95 | #sensors | #gu | 기간")
for y in sorted(by):
    v=sorted(by[y]); n=len(v)
    print(f"  {y} | {n:,} | {v[int(n*.05)]:.0f}/{v[n//2]:.0f}/{v[int(n*.95)]:.0f} dB | {len(sensors[y])} | {len(gus[y])}")
# 날짜 범위
dates=[]
for row in csv.DictReader(open(panel,encoding="utf-8-sig")):
    dates.append(row["date"])
print(f"\n전체 날짜 범위: {min(dates)} ~ {max(dates)} | 총 행 {len(dates):,}")
print("=== DONE ===")
