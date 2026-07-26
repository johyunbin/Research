# -*- coding: utf-8 -*-
import os, zipfile, csv, io, sys

base = r"C:\Users\wh850\AppData\Local\Temp\sdot"
print("=== base dir listing ===")
try:
    for f in os.listdir(base):
        p = os.path.join(base, f)
        print(f"  {f}  ({os.path.getsize(p):,} bytes)")
except Exception as e:
    print("  ERR listing:", e); sys.exit(1)

def decode(b):
    for enc in ("cp949", "euc-kr", "utf-8-sig", "utf-8"):
        try:
            return b.decode(enc), enc
        except Exception:
            continue
    return b.decode("utf-8", "replace"), "utf-8-replace"

# pick a zip
zpath = None
for cand in ("sdot2022.zip", "sdot2020.zip"):
    if os.path.exists(os.path.join(base, cand)):
        zpath = os.path.join(base, cand); break
if not zpath:
    zs = [f for f in os.listdir(base) if f.lower().endswith(".zip")]
    if zs: zpath = os.path.join(base, zs[0])
print(f"\n=== using zip: {os.path.basename(zpath) if zpath else None} ===")
if not zpath: sys.exit(1)

z = zipfile.ZipFile(zpath)
names = z.namelist()
csvs = [n for n in names if n.lower().endswith(".csv")]
print(f"  members total={len(names)}, csv={len(csvs)}")
# decode member names for display
for n in names[:6]:
    nm, _ = decode(n.encode("cp437")) if False else (n, "")
    print("   member:", n)

# choose a representative csv (mid-list = mid-year)
member = csvs[len(csvs)//2] if csvs else None
print(f"\n=== reading member: {member} ===")
raw = z.read(member)
text, enc = decode(raw)
print(f"  decoded with: {enc}, bytes={len(raw):,}")

rdr = list(csv.reader(io.StringIO(text)))
print(f"  rows (incl header) = {len(rdr):,}")
header = rdr[0]
print("\n--- header (index : name) ---")
for i, h in enumerate(header):
    print(f"  [{i:2d}] {h}")

# locate key columns
def find(colpart):
    for i, h in enumerate(header):
        if colpart in h: return i
    return -1
i_ser = find("시리얼")
i_noise = find("소음")
i_reg = find("등록")
i_tx = find("전송")
i_gubun = find("구분")
print(f"\n  col idx -> 시리얼={i_ser} 소음={i_noise} 등록일자={i_reg} 전송시간={i_tx} 구분={i_gubun}")

print("\n--- first 5 data rows (key cols) ---")
for r in rdr[1:6]:
    def g(i): return r[i] if (0 <= i < len(r)) else "?"
    print(f"   시리얼={g(i_ser)} | 구분={g(i_gubun)} | 소음={g(i_noise)} | 등록일자={g(i_reg)} | 전송시간={g(i_tx)}")

# noise stats + missingness (stream)
tot = 0; nonempty = 0; vals = []
regfmts = {}
for r in rdr[1:]:
    tot += 1
    v = r[i_noise].strip() if (0 <= i_noise < len(r)) else ""
    if v != "":
        nonempty += 1
        try: vals.append(float(v))
        except: pass
    if 0 <= i_reg < len(r):
        rv = r[i_reg].strip()
        # bucket by format signature
        sig = "".join("D" if c.isdigit() else c for c in rv)[:16]
        regfmts[sig] = regfmts.get(sig, 0) + 1

print(f"\n--- 소음(dB) ---")
print(f"  total data rows = {tot:,}")
print(f"  non-empty noise = {nonempty:,} ({100*nonempty/max(tot,1):.1f}%)")
if vals:
    vals.sort()
    n = len(vals)
    mean = sum(vals)/n
    print(f"  parsed floats   = {n:,}")
    print(f"  min={vals[0]:.1f}  p05={vals[int(n*0.05)]:.1f}  median={vals[n//2]:.1f}  mean={mean:.1f}  p95={vals[int(n*0.95)]:.1f}  max={vals[-1]:.1f}")

print(f"\n--- 등록일자 format signatures (top) ---")
for sig, c in sorted(regfmts.items(), key=lambda x:-x[1])[:5]:
    print(f"   '{sig}'  x{c:,}")

# distinct sensors in this file
ser = set()
for r in rdr[1:]:
    if 0 <= i_ser < len(r): ser.add(r[i_ser])
print(f"\n  distinct 시리얼 (sensors) in this weekly file = {len(ser):,}")

# location xlsx peek
print("\n=== location.xlsx peek ===")
loc = os.path.join(base, "location.xlsx")
if os.path.exists(loc):
    try:
        import openpyxl
        wb = openpyxl.load_workbook(loc, read_only=True)
        ws = wb.active
        rows = ws.iter_rows(values_only=True)
        hdr = next(rows)
        print("  header:", hdr)
        for i, row in enumerate(rows):
            if i >= 3: break
            print("   row:", row)
        cnt = 3 + sum(1 for _ in rows)
        print(f"  ~total rows (incl shown) approx = {cnt+1}")
    except Exception as e:
        print("  (openpyxl unavailable or err:", e, ")")
else:
    print("  location.xlsx not found")
print("\n=== DONE ===")
