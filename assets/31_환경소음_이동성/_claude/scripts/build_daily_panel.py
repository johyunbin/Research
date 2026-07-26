# -*- coding: utf-8 -*-
# S-DoT 일(日) 패널 빌드 (within-sensor 분석용 결과변수)
# 센서×일: Leq24/주간(06-21)/야간(22-05) energy-avg + 좌표·구 + QC. 분석창 2020-04~2023.
import os, zipfile, io, csv, math
from collections import defaultdict
import openpyxl

build_dir = r"C:\Users\wh850\AppData\Local\Temp\sdot_build"
meta = r"T:\00_BACKUP\01_한양대학교\00_연구실\00_프로젝트\★_논문화 프로젝트\31_환경소음_이동성\data\sdot_meta\location.xlsx"
MIN_HOURS = 12          # 일 최소 관측시간
QC_LO, QC_HI = 20, 95   # 물리적 dB 범위

def decode(b):
    for e in ("cp949","euc-kr","utf-8-sig","utf-8"):
        try: return b.decode(e)
        except: pass
    return b.decode("utf-8","replace")
def cidx(h,p):
    for i,x in enumerate(h):
        if p in (x or ""): return i
    return -1
def noise_idx(h):
    # 2023 스키마는 소음을 최대/평균/최소로 분할 → 평균 우선, 최대/최소 회피 (2020-22 단일 '소음(dB)'는 그대로)
    for i,c in enumerate(h):
        if "소음" in (c or "") and "평균" in (c or ""): return i
    for i,c in enumerate(h):
        if "소음" in (c or "") and "최대" not in (c or "") and "최소" not in (c or ""): return i
    return -1
def gu_of(a):
    for t in (a or "").split():
        if t.endswith("구"): return t
    return ""
def leq(s,n): return round(10*math.log10(s/n),2) if n>0 else ""

# 좌표·구
wb=openpyxl.load_workbook(meta,read_only=True)
rows=list(wb.active.iter_rows(values_only=True)); hdr=[str(x) if x is not None else "" for x in rows[0]]
lat_i=cidx(hdr,"위도"); lon_i=cidx(hdr,"경도"); ad_i=cidx(hdr,"주소")
ser2meta={}
for r in rows[1:]:
    if r[1]:
        try: ser2meta[str(r[1]).strip()]=(float(r[lat_i]),float(r[lon_i]),gu_of(str(r[ad_i]) if ad_i<len(r) and r[ad_i] else ""))
        except: pass
print("meta sensors:", len(ser2meta))

def build_year(zpath, writer):
    acc=defaultdict(lambda: [0.0,0,0.0,0,0.0,0])  # (ser,date)->[s24,n24,sday,nday,snt,nnt]
    z=zipfile.ZipFile(zpath)
    bad=0; firsthdr=True
    for n in [x for x in z.namelist() if x.lower().endswith(".csv")]:
        rdr=csv.reader(io.StringIO(decode(z.read(n)))); h=next(rdr)
        si=cidx(h,"시리얼"); ni=noise_idx(h); ri=cidx(h,"등록")
        if firsthdr: print(f"   {os.path.basename(zpath)}: cols={len(h)} 소음컬럼={h[ni]!r}(idx{ni}) 시리얼idx={si} 등록idx={ri}"); firsthdr=False
        for r in rdr:
            if si<len(r) and ni<len(r) and ri<len(r):
                v=r[ni].strip()
                if not v: continue
                try: fv=float(v)
                except: continue
                if fv<=0 or fv>140: bad+=1; continue
                ts=r[ri].strip()
                if not ts: continue
                dp=ts.split(); p=dp[0].replace(".","-").split("-")
                if len(p)!=3 or not all(x.isdigit() for x in p): continue
                ds=f"{int(p[0]):04d}-{int(p[1]):02d}-{int(p[2]):02d}"
                hh=None
                if len(dp)>1:
                    try: hh=int(dp[1].split(":")[0])
                    except: hh=None
                lin=10**(fv/10.0); a=acc[(r[si].strip(),ds)]
                a[0]+=lin; a[1]+=1
                if hh is not None:
                    if 6<=hh<=21: a[2]+=lin; a[3]+=1
                    else: a[4]+=lin; a[5]+=1
    tot=len(acc); kept=0
    for (ser,ds),a in acc.items():
        if a[1]<MIN_HOURS: continue
        l24=10*math.log10(a[0]/a[1])
        if l24<QC_LO or l24>QC_HI: continue
        m=ser2meta.get(ser,("","",""))
        writer.writerow([ser,m[0],m[1],m[2],ds,a[1],leq(a[0],a[1]),leq(a[2],a[3]),leq(a[4],a[5])])
        kept+=1
    if bad: print(f"   skipped garbage(소음<=0 or >140dB): {bad:,}")
    return tot,kept

out=os.path.join(build_dir,"sdot_daily_panel.csv")
with open(out,"w",newline="",encoding="utf-8-sig") as f:
    w=csv.writer(f)
    w.writerow(["serial","lat","lon","gu","date","n_hours","Leq24","Leq_day","Leq_night"])
    for y in (2020,2021,2022,2023):
        zp=os.path.join(build_dir,f"sdot_{y}.zip")
        if not os.path.exists(zp): print(f"{y}: zip 없음 → skip"); continue
        tot,kept=build_year(zp,w)
        print(f"{y}: sensor-days={tot:,}, kept(>={MIN_HOURS}h & QC)={kept:,}")
print("-> sdot_daily_panel.csv")
print("=== DONE ===")
