# -*- coding: utf-8 -*-
# 서울 146개 환경소음 측정지점 + 2024 연 LAeq(주/야, 분기 에너지평균) 추출 → 지오코딩/검증용
import os, csv, math

tmp = r"C:\Users\wh850\AppData\Local\Temp"
sdot_dir = os.path.join(tmp, "sdot")
vp = os.path.join(tmp, "values_15065396.csv")

def load_csv(path):
    for enc in ("utf-8-sig","cp949","euc-kr","utf-8"):
        try:
            with open(path, encoding=enc, newline="") as f:
                return list(csv.reader(f))
        except Exception: continue
    return None
def ci(h,p):
    for i,x in enumerate(h):
        if p in (x or ""): return i
    return -1
def eavg(vals):
    vv=[v for v in vals if v is not None]
    return 10*math.log10(sum(10**(v/10.0) for v in vv)/len(vv)) if vv else None

rows = load_csv(vp); h = rows[0]
c_city=ci(h,"도시"); c_pt=ci(h,"측정지점"); c_reg=ci(h,"지역")
c_dm=ci(h,"주간평균"); c_nm=ci(h,"야간평균")
seoul=[r for r in rows[1:] if c_city<len(r) and "서울" in r[c_city]]

from collections import defaultdict
day=defaultdict(list); ngt=defaultdict(list); zone={}
for r in seoul:
    pt=r[c_pt].strip()
    zone[pt]=r[c_reg].strip() if c_reg<len(r) else ""
    try: day[pt].append(float(r[c_dm]))
    except: pass
    try: ngt[pt].append(float(r[c_nm]))
    except: pass

out=os.path.join(sdot_dir,"reference_seoul_146.csv")
with open(out,"w",newline="",encoding="utf-8-sig") as f:
    w=csv.writer(f)
    w.writerow(["측정지점","지역","n_q","Leq_day_2024","Leq_night_2024"])
    for pt in sorted(day):
        dd=eavg(day[pt]); nn=eavg(ngt[pt])
        w.writerow([pt, zone.get(pt,""), len(day[pt]),
                    round(dd,1) if dd is not None else "", round(nn,1) if nn is not None else ""])
print("points:", len(day), "-> reference_seoul_146.csv")
dv=sorted(eavg(day[p]) for p in day if eavg(day[p]) is not None)
n=len(dv)
print(f"day LAeq 2024: min {dv[0]:.0f} / median {dv[n//2]:.0f} / max {dv[-1]:.0f} dB")

# 이름 목록(지오코딩 에이전트 전달용) 별도 저장
names=sorted(day)
with open(os.path.join(sdot_dir,"ref146_names.txt"),"w",encoding="utf-8") as f:
    for nm in names: f.write(nm+"\n")
print("names -> ref146_names.txt")
print("sample names:", names[:10])
print("=== DONE ===")
