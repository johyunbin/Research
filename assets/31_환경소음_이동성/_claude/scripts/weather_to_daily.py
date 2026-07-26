# -*- coding: utf-8 -*-
# Open-Meteo ERA5 재분석 일별 기상(서울 종관108 좌표) JSON -> CSV.
# 변수: 평균/최고/최저기온(C), 강수합(mm), 최대풍속(km/h->m/s), 우세풍향. 소음 교란 통제용 공변량.
# 무인증·재현가능(ERA5). KMA ASOS 공식자료는 사용자 API키 보유 시 대체 가능.
import json, csv, os

ROOT = r"T:\00_BACKUP\01_한양대학교\00_연구실\00_프로젝트\★_논문화 프로젝트\31_환경소음_이동성"
SRC = r"C:\Users\wh850\AppData\Local\Temp\weather_openmeteo.json"
OUT = os.path.join(ROOT, r"data\processed\weather_seoul_daily_2020-2023.csv")

j = json.load(open(SRC, encoding="utf-8"))
D = j["daily"]
t = D["time"]
n = len(t)

def g(k):
    return D.get(k, [None] * n)

tmean = g("temperature_2m_mean"); tmax = g("temperature_2m_max"); tmin = g("temperature_2m_min")
prcp = g("precipitation_sum"); wmax = g("wind_speed_10m_max"); wdir = g("wind_direction_10m_dominant")

with open(OUT, "w", newline="", encoding="utf-8-sig") as f:
    w = csv.writer(f)
    w.writerow(["date", "temp_mean", "temp_max", "temp_min", "precip_mm",
                "wind_max_ms", "wind_dir", "rain_flag"])
    for i in range(n):
        wms = round(wmax[i] / 3.6, 2) if wmax[i] is not None else ""
        rain = 1 if (prcp[i] is not None and prcp[i] > 0) else 0
        w.writerow([t[i], tmean[i], tmax[i], tmin[i], prcp[i], wms, wdir[i], rain])

print(f"일수: {n} ({t[0]} ~ {t[-1]})")
print(f"-> {OUT}")
# sanity
import statistics
tm = [x for x in tmean if x is not None]
pr = [x for x in prcp if x is not None]
rain_days = sum(1 for x in prcp if x and x > 0)
print(f"기온 평균 {statistics.mean(tm):.1f}C (min {min(tm):.1f} / max {max(tm):.1f})")
print(f"연 강수합 평균 {sum(pr)/4:.0f}mm | 강수일 {rain_days}/{n} ({rain_days/n*100:.0f}%)")
print(f"결측: tmean={sum(1 for x in tmean if x is None)} precip={sum(1 for x in prcp if x is None)}")
