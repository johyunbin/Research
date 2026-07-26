# -*- coding: utf-8 -*-
# 받아둔 기준데이터 점검: 15065396(값) 지역 granularity + 15093409(좌표) 서울 + S-DoT 구 분포
import os, csv
import openpyxl

tmp = r"C:\Users\wh850\AppData\Local\Temp"
sdot_dir = os.path.join(tmp, "sdot")

def load_csv(path):
    for enc in ("cp949","euc-kr","utf-8-sig","utf-8"):
        try:
            with open(path, encoding=enc, newline="") as f:
                return list(csv.reader(f)), enc
        except Exception:
            continue
    return None, None
def ci(h, part):
    for i,x in enumerate(h):
        if part in (x or ""): return i
    return -1
def gu_of(addr):
    for t in (addr or "").split():
        if t.endswith("구"): return t
    return None

# --- values 15065396 ---
vp = os.path.join(tmp, "values_15065396.csv")
print("=== values_15065396.csv exists:", os.path.exists(vp), "===")
if os.path.exists(vp):
    rows, enc = load_csv(vp)
    h = rows[0]; print("enc", enc, "| rows", len(rows)); print("header:", h)
    c_city=ci(h,"도시"); c_pt=ci(h,"측정지점"); c_reg=ci(h,"지역"); c_yr=ci(h,"측정연도"); c_q=ci(h,"분기")
    c_dm=ci(h,"주간평균"); c_nm=ci(h,"야간평균")
    seoul=[r for r in rows[1:] if c_city>=0 and c_city<len(r) and "서울" in r[c_city]]
    print("seoul rows:", len(seoul))
    pts=set(r[c_pt] for r in seoul if c_pt<len(r))
    regs=[r[c_reg] for r in seoul if c_reg<len(r) and r[c_reg]]
    regset=set(regs)
    print("distinct 측정지점:", len(pts), "| distinct 지역:", len(regset))
    print("지역 값 예시:", list(regset)[:25])
    print("측정지점 예시:", list(pts)[:12])
    print("측정연도 분포:", sorted(set(r[c_yr] for r in seoul if c_yr<len(r))))
    print("sample seoul rows (지점|지역|연도|분기|주간평균|야간평균):")
    for r in seoul[:6]:
        g=lambda i: r[i] if i>=0 and i<len(r) else "?"
        print("   ", g(c_pt),"|",g(c_reg),"|",g(c_yr),"|",g(c_q),"|",g(c_dm),"|",g(c_nm))

# --- coords 15093409 ---
cp = os.path.join(tmp, "coords_15093409.csv")
print("\n=== coords_15093409.csv exists:", os.path.exists(cp), "===")
if os.path.exists(cp):
    rows, enc = load_csv(cp)
    h=rows[0]; print("enc", enc, "| rows", len(rows)); print("header:", h)
    c_pt=ci(h,"측정지점"); c_ad=ci(h,"주소"); c_la=ci(h,"위도"); c_lo=ci(h,"경도")
    seoul=[r for r in rows[1:] if c_ad>=0 and c_ad<len(r) and "서울" in r[c_ad]]
    print("seoul coord rows:", len(seoul))
    for r in seoul[:12]:
        print("   ", r[c_pt] if c_pt<len(r) else "?", "|", r[c_la] if c_la<len(r) else "?", r[c_lo] if c_lo<len(r) else "?", "|", r[c_ad] if c_ad<len(r) else "?")

# --- S-DoT 구 분포 (location.xlsx 주소에서) ---
print("\n=== S-DoT 구 분포 (location.xlsx) ===")
wb=openpyxl.load_workbook(os.path.join(sdot_dir,"location.xlsx"),read_only=True)
rws=list(wb.active.iter_rows(values_only=True)); hdr=[str(x) if x is not None else "" for x in rws[0]]
a_i=ci(hdr,"주소")
from collections import Counter
cnt=Counter()
for r in rws[1:]:
    if a_i<len(r):
        g=gu_of(str(r[a_i]) if r[a_i] else "")
        if g: cnt[g]+=1
print("S-DoT 구 수:", len(cnt), "| 총센서(구추출):", sum(cnt.values()))
print("구별 센서수:", dict(sorted(cnt.items(), key=lambda x:-x[1])))
print("\n=== DONE ===")
