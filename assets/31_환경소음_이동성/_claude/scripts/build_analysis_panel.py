# -*- coding: utf-8 -*-
# Phase 1-E 분석패널 결합: 결과(ΔLAeq) + 이동량 dose(생활인구·지하철) + 거리두기 + 기상 + 달력.
# within-sensor 설계: 모델은 (1|sensor) 또는 센서FE; 기술/지도용 ΔLAeq = Leq - 센서 post-lift 기준평균.
# dose = within-dong 상대변화(동 post-lift 기준 대비). 크로스워크 dose_key(강북 relabel·상일 병합).
import os
import numpy as np
import pandas as pd

ROOT = r"T:\00_BACKUP\01_한양대학교\00_연구실\00_프로젝트\★_논문화 프로젝트\31_환경소음_이동성"
P = os.path.join(ROOT, "data", "processed")
OUT = os.path.join(P, "analysis_panel.csv")
BASE0, BASE1 = "2022-07-01", "2023-12-31"   # post-lift 정상기 기준창 (zero-dose baseline)

# 검증된 코드 크로스워크 (2026-06-21): 강북 6동 livingpop->vuski, 강동 상일 분할 병합
XWALK = {"11305590": "11305595", "11305600": "11305603", "11305606": "11305608",
         "11305610": "11305615", "11305620": "11305625", "11305630": "11305635",
         "11740525": "11740520", "11740526": "11740520"}
def dose_key(s):
    return s.map(lambda c: XWALK.get(c, c))

# 한국 공휴일 2020-2023 (대체·임시공휴일 포함; 소음·이동 교란 통제 공변량)
HOLIDAYS = set(pd.to_datetime([
    # 2020
    "2020-01-01","2020-01-24","2020-01-25","2020-01-26","2020-01-27","2020-03-01",
    "2020-04-15","2020-04-30","2020-05-05","2020-06-06","2020-08-15","2020-08-17",
    "2020-09-30","2020-10-01","2020-10-02","2020-10-03","2020-10-09","2020-12-25",
    # 2021
    "2021-01-01","2021-02-11","2021-02-12","2021-02-13","2021-03-01","2021-05-05",
    "2021-05-19","2021-06-06","2021-08-15","2021-08-16","2021-09-20","2021-09-21",
    "2021-09-22","2021-10-03","2021-10-04","2021-10-09","2021-10-11","2021-12-25",
    # 2022
    "2022-01-01","2022-01-31","2022-02-01","2022-02-02","2022-03-01","2022-03-09",
    "2022-05-05","2022-05-08","2022-06-01","2022-06-06","2022-08-15","2022-09-09",
    "2022-09-10","2022-09-11","2022-09-12","2022-10-03","2022-10-09","2022-10-10","2022-12-25",
    # 2023
    "2023-01-01","2023-01-21","2023-01-22","2023-01-23","2023-01-24","2023-03-01",
    "2023-05-05","2023-05-27","2023-05-29","2023-06-06","2023-08-15","2023-09-28",
    "2023-09-29","2023-09-30","2023-10-02","2023-10-03","2023-10-09","2023-12-25",
]))
SEASON = {12: "DJF", 1: "DJF", 2: "DJF", 3: "MAM", 4: "MAM", 5: "MAM",
          6: "JJA", 7: "JJA", 8: "JJA", 9: "SON", 10: "SON", 11: "SON"}

print("[1] 결과 패널 로드...")
df = pd.read_csv(os.path.join(P, "sdot_daily_panel_2020-2023.csv"), dtype={"serial": str})
df["date"] = pd.to_datetime(df["date"])
n0 = len(df)
print(f"    sensor-days={n0:,}  sensors={df['serial'].nunique()}")

print("[2] 센서->행정동 결합...")
sd = pd.read_csv(os.path.join(P, "sensor_dong_map.csv"), dtype={"serial": str, "adm_cd": str})
sd = sd[sd["matched"] == 1][["serial", "adm_cd", "adm_nm"]]
df = df.merge(sd, on="serial", how="inner")
df["dose_key"] = dose_key(df["adm_cd"])
print(f"    동 결합 후 sensor-days={len(df):,} (동 미배정 제외 {n0-len(df):,})  동={df['dose_key'].nunique()}")

print("[3] 생활인구 dose 결합 (dose_key+date SUM)...")
lp = pd.read_csv(os.path.join(P, "livingpop_dong_daily_2020-2023.csv"), dtype={"adm_cd": str})
lp["date"] = pd.to_datetime(lp["date"])
lp["dose_key"] = dose_key(lp["adm_cd"])
lp_agg = lp.groupby(["dose_key", "date"], as_index=False)[["lp_mean", "lp_day", "lp_night"]].sum()
df = df.merge(lp_agg, on=["dose_key", "date"], how="left")
# 동 post-lift 기준 -> 상대 dose
post_lp = lp_agg[(lp_agg["date"] >= BASE0) & (lp_agg["date"] <= BASE1)]
db = post_lp.groupby("dose_key")[["lp_mean", "lp_day", "lp_night"]].mean()
db.columns = ["base_lp_mean", "base_lp_day", "base_lp_night"]
df = df.merge(db, on="dose_key", how="left")
for k in ["mean", "day", "night"]:
    df[f"lp_{k}_rel"] = df[f"lp_{k}"] / df[f"base_lp_{k}"]
    df[f"lp_{k}_logrel"] = np.log(df[f"lp_{k}_rel"].where(df[f"lp_{k}_rel"] > 0))
print(f"    생활인구 결측 sensor-days: {df['lp_day'].isna().sum():,} ({df['lp_day'].isna().mean()*100:.2f}%)")

print("[4] 지하철·거리두기·기상 결합 (date)...")
sub = pd.read_csv(os.path.join(P, "subway_daily_seoul_2020-2023.csv")); sub["date"] = pd.to_datetime(sub["date"])
sub_base = sub[(sub["date"] >= BASE0) & (sub["date"] <= BASE1)]["subway_total"].mean()
sub["subway_rel"] = sub["subway_total"] / sub_base
df = df.merge(sub, on="date", how="left")

dis = pd.read_csv(os.path.join(P, "distancing_daily_2020-2023.csv")); dis["date"] = pd.to_datetime(dis["date"])
df = df.merge(dis[["date", "stringency", "close_hour", "gather_limit", "any_distancing", "phase_label"]], on="date", how="left")

wx = pd.read_csv(os.path.join(P, "weather_seoul_daily_2020-2023.csv")); wx["date"] = pd.to_datetime(wx["date"])
df = df.merge(wx[["date", "temp_mean", "temp_max", "temp_min", "precip_mm", "wind_max_ms", "rain_flag"]], on="date", how="left")

print("[5] 센서 baseline ΔLAeq + 달력...")
post = df[(df["date"] >= BASE0) & (df["date"] <= BASE1)]
sb = post.groupby("serial")[["Leq24", "Leq_day", "Leq_night"]].mean()
sb.columns = ["base_Leq24", "base_Leq_day", "base_Leq_night"]
sb_n = post.groupby("serial").size().rename("base_n")
df = df.merge(sb, on="serial", how="left").merge(sb_n, on="serial", how="left")
for k in ["Leq24", "Leq_day", "Leq_night"]:
    df[f"d{k}"] = df[k] - df[f"base_{k}"]
df["dow"] = df["date"].dt.dayofweek
df["is_weekend"] = (df["dow"] >= 5).astype(int)
df["month"] = df["date"].dt.month
df["year"] = df["date"].dt.year
df["season"] = df["month"].map(SEASON)
df["holiday"] = df["date"].isin(HOLIDAYS).astype(int)

cols = ["serial", "adm_cd", "dose_key", "adm_nm", "gu", "lat", "lon", "date", "year", "month", "season",
        "dow", "is_weekend", "holiday", "n_hours", "Leq24", "Leq_day", "Leq_night",
        "base_Leq24", "base_n", "dLeq24", "dLeq_day", "dLeq_night",
        "lp_mean", "lp_day", "lp_night", "lp_day_rel", "lp_day_logrel", "lp_mean_rel", "lp_mean_logrel",
        "lp_night_rel", "lp_night_logrel",
        "subway_total", "subway_rel", "stringency", "close_hour", "gather_limit", "any_distancing",
        "temp_mean", "temp_max", "temp_min", "precip_mm", "wind_max_ms", "rain_flag"]
out = df[cols].sort_values(["serial", "date"])
out.to_csv(OUT, index=False, encoding="utf-8-sig")

print(f"\n=== 패널 완성: {len(out):,} sensor-days ===")
print(f"센서 {out['serial'].nunique()} | 동 {out['dose_key'].nunique()} | 기간 {out['date'].min().date()}~{out['date'].max().date()}")
print(f"ΔLAeq 가용(baseline 있는 센서): {out['dLeq24'].notna().sum():,} ({out['dLeq24'].notna().mean()*100:.1f}%)")
print(f"dose 가용(lp_day_rel): {out['lp_day_rel'].notna().sum():,} ({out['lp_day_rel'].notna().mean()*100:.1f}%)")
print(f"-> {OUT}  ({os.path.getsize(OUT)/1e6:.0f}MB)")
print("\n동시가용(ΔLAeq & dose & 기상) 행:", (out['dLeq24'].notna() & out['lp_day_rel'].notna() & out['temp_mean'].notna()).sum())
