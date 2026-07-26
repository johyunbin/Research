# -*- coding: utf-8 -*-
# 센서 좌표 -> 행정동 공간조인 (point-in-polygon, 순수 Python ray-casting).
# 경계 = vuski/admdongkor ver20220101. 조인키 adm_cd = adm_cd2[:8] (= 생활인구 OA-14991 행정동코드).
# location.xlsx: col1=보정시리얼, col2=주소, col4=위도, col5=경도.
# 출력: data/processed/sensor_dong_map.csv (serial,lat,lon,addr_gu,adm_cd,adm_nm,dong_gu,matched)
import sys, os, json, csv
import openpyxl

ROOT = r"T:\00_BACKUP\01_한양대학교\00_연구실\00_프로젝트\★_논문화 프로젝트\31_환경소음_이동성"
GEO  = r"C:\Users\wh850\AppData\Local\Temp\admdong_ver20220101.geojson"
META = os.path.join(ROOT, r"data\sdot_meta\location.xlsx")
OUT  = os.path.join(ROOT, r"data\processed\sensor_dong_map.csv")
SEO_GEO_OUT = os.path.join(ROOT, r"data\reference\admdong_seoul_ver20220101.geojson")  # Phase2 지도 재사용

def ring_contains(x, y, ring):
    inside = False; n = len(ring); j = n - 1
    for i in range(n):
        xi, yi = ring[i][0], ring[i][1]
        xj, yj = ring[j][0], ring[j][1]
        if ((yi > y) != (yj > y)) and (x < (xj - xi) * (y - yi) / (yj - yi + 1e-18) + xi):
            inside = not inside
        j = i
    return inside

def poly_contains(x, y, poly):  # poly = [exterior, hole1, ...]
    if not poly or not ring_contains(x, y, poly[0]):
        return False
    for hole in poly[1:]:
        if ring_contains(x, y, hole):
            return False
    return True

def gu_of(addr):
    for t in (addr or "").split():
        if t.endswith("구"):
            return t
    return ""

# --- 서울 동 경계 로드 ---
g = json.load(open(GEO, encoding="utf-8"))
dongs = []  # (adm_cd, adm_nm, gu, [polys], bbox)
seoul_feats = []
for f in g["features"]:
    p = f["properties"]
    if not str(p.get("adm_cd2", "")).startswith("11"):
        continue
    seoul_feats.append(f)
    adm_cd = str(p["adm_cd2"])[:8]
    geom = f["geometry"]
    polys = geom["coordinates"] if geom["type"] == "MultiPolygon" else [geom["coordinates"]]
    xs = [pt[0] for poly in polys for ring in poly for pt in ring]
    ys = [pt[1] for poly in polys for ring in poly for pt in ring]
    bbox = (min(xs), min(ys), max(xs), max(ys))
    dongs.append((adm_cd, p["adm_nm"], p.get("sggnm", ""), polys, bbox))
print(f"서울 행정동 폴리곤: {len(dongs)}")

# 서울 경계만 저장(Phase2 재사용)
json.dump({"type": "FeatureCollection", "features": seoul_feats},
          open(SEO_GEO_OUT, "w", encoding="utf-8"), ensure_ascii=False)

# --- 센서 로드 ---
wb = openpyxl.load_workbook(META, read_only=True)
rows = list(wb.active.iter_rows(values_only=True))
sensors = []
for r in rows[1:]:
    if not r[1]:
        continue
    try:
        lat = float(r[4]); lon = float(r[5])
    except (TypeError, ValueError):
        continue
    addr = str(r[2]) if r[2] else ""
    sensors.append((str(r[1]).strip(), lat, lon, gu_of(addr)))
print(f"센서(좌표 유효): {len(sensors)}")

# --- point-in-polygon (bbox 프리필터) ---
matched = 0; gu_ok = 0; gu_bad = []
with open(OUT, "w", newline="", encoding="utf-8-sig") as fo:
    w = csv.writer(fo)
    w.writerow(["serial", "lat", "lon", "addr_gu", "adm_cd", "adm_nm", "dong_gu", "matched"])
    for ser, lat, lon, agu in sensors:
        x, y = lon, lat
        hit = None
        for adm_cd, adm_nm, dgu, polys, bbox in dongs:
            if x < bbox[0] or x > bbox[2] or y < bbox[1] or y > bbox[3]:
                continue
            for poly in polys:
                if poly_contains(x, y, poly):
                    hit = (adm_cd, adm_nm, dgu); break
            if hit:
                break
        if hit:
            matched += 1
            if agu and hit[2] and agu == hit[2]:
                gu_ok += 1
            elif agu and hit[2]:
                gu_bad.append((ser, agu, hit[2], adm_nm if hit else ""))
            w.writerow([ser, lat, lon, agu, hit[0], hit[1], hit[2], 1])
        else:
            w.writerow([ser, lat, lon, agu, "", "", "", 0])

print(f"\n매칭: {matched}/{len(sensors)} ({matched/len(sensors)*100:.1f}%)")
print(f"구 일치(주소구 vs 동의구): {gu_ok}/{matched} ({gu_ok/matched*100:.1f}%)")
if gu_bad:
    print(f"구 불일치 {len(gu_bad)}건 (경계 인접/좌표오류 가능) 샘플:")
    for s in gu_bad[:8]:
        print("  ", s)
print(f"\n-> {OUT}")
print(f"-> {SEO_GEO_OUT} (서울 {len(seoul_feats)}동 경계, Phase2 지도용)")
