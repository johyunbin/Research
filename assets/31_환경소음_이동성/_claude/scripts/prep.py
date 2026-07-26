# -*- coding: utf-8 -*-
# S-DoT 검증 스파이크 준비: (1) 시리얼<->좌표 조인 해법 (2) 일집계 가능성 점검
import os, zipfile, csv, io, sys
from collections import defaultdict

base = r"C:\Users\wh850\AppData\Local\Temp\sdot"

def decode(b):
    for enc in ("cp949","euc-kr","utf-8-sig","utf-8"):
        try: return b.decode(enc)
        except: pass
    return b.decode("utf-8","replace")

def read_member(zpath, prefer):
    z = zipfile.ZipFile(zpath)
    csvs = [n for n in z.namelist() if n.lower().endswith(".csv")]
    member = None
    for sub in prefer:
        for n in csvs:
            if sub in n: member = n; break
        if member: break
    if not member: member = csvs[len(csvs)//2]
    return member, list(csv.reader(io.StringIO(decode(z.read(member)))))

def cidx(header, part):
    for i,h in enumerate(header):
        if part in h: return i
    return -1

def fam(s):
    s=(s or "").strip()
    if s.startswith("V02"): return "V02x(2020형)"
    if s.startswith("OC3"): return "OC3x(2022형)"
    if not s: return "(빈)"
    return "기타:"+s[:4]

# ---- 1) location.xlsx : serial 컬럼들 + 좌표 ----
print("=== location.xlsx ===")
import openpyxl
wb = openpyxl.load_workbook(os.path.join(base,"location.xlsx"), read_only=True)
ws = wb.active
rows = list(ws.iter_rows(values_only=True))
hdr = [str(x) if x is not None else "" for x in rows[0]]
data = rows[1:]
print("header:", hdr)
print("total rows:", len(data))
serial_cols = [i for i,h in enumerate(hdr) if "시리얼" in h]
lat_i = cidx(hdr,"위도"); lon_i = cidx(hdr,"경도")
print("serial cols:", [(i,hdr[i]) for i in serial_cols], "| lat,lon idx:", lat_i, lon_i)
loc_sets = {}
for i in serial_cols:
    vals = [str(r[i]).strip() for r in data if i < len(r) and r[i] not in (None,"")]
    fmt = defaultdict(int)
    for v in vals: fmt[fam(v)] += 1
    loc_sets[i] = set(vals)
    print(f"  col[{i}] {hdr[i]}: non-empty={len(vals)} formats={dict(fmt)} ex={vals[:2]}")

# ---- 2) S-DoT 연도별 시리얼 집합 ----
def serials(zname, prefer):
    m, rws = read_member(os.path.join(base,zname), prefer)
    h = rws[0]; si = cidx(h,"시리얼")
    return m, set(r[si].strip() for r in rws[1:] if si<len(r) and r[si].strip())
m22, s22 = serials("sdot2022.zip", ["2022.07.11","2022.07","2022.06"])
m20, s20 = serials("sdot2020.zip", ["2020.11","2020.10","2020.12"])
print(f"\n2022 file={m22}: sensors={len(s22)} ex={list(s22)[:2]}")
print(f"2020 file={m20}: sensors={len(s20)} ex={list(s20)[:2]}")

# ---- 3) 조인 테스트: 각 연도 시리얼이 어느 location 컬럼과 매칭되나 ----
print("\n=== JOIN test (S-DoT serials ∩ location columns) ===")
for i in serial_cols:
    print(f"  loc[{i}] {hdr[i]}: 2022∩={len(s22 & loc_sets[i])}/{len(s22)}  2020∩={len(s20 & loc_sets[i])}/{len(s20)}")

# ---- 4) 일집계 가능성 (2022 주간 파일) ----
print("\n=== daily aggregation test (2022 weekly) ===")
m, rws = read_member(os.path.join(base,"sdot2022.zip"), ["2022.07.11","2022.07"])
h = rws[0]; si=cidx(h,"시리얼"); ni=cidx(h,"소음"); ri=cidx(h,"등록")
def to_date(s):
    s=(s or "").strip()
    s=s.split()[0] if s else ""
    s=s.replace(".","-")
    p=s.split("-")
    return f"{int(p[0]):04d}-{int(p[1]):02d}-{int(p[2]):02d}" if len(p)==3 and all(x.isdigit() for x in p) else s
agg=defaultdict(list)
for r in rws[1:]:
    if max(si,ni,ri) < len(r):
        v=r[ni].strip()
        if not v: continue
        try: fv=float(v)
        except: continue
        agg[(r[si].strip(), to_date(r[ri]))].append(fv)
counts=sorted(len(v) for v in agg.values()); n=len(counts)
print(f"  (sensor,date) groups={n}")
if n:
    print(f"  hours/day: min={counts[0]} median={counts[n//2]} max={counts[-1]}")
    for thr in (24,18,12,6):
        c=sum(1 for x in counts if x>=thr)
        print(f"   >= {thr} hrs/day: {c} ({100*c/n:.0f}%)")
    import statistics
    for (ser,d),vals in list(agg.items())[:5]:
        print(f"   {ser} {d}: n={len(vals)} dailyMean={statistics.mean(vals):.1f}dB")
print("\n=== DONE ===")
