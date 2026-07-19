# -*- coding: utf-8 -*-
# figs_v2 준비: 550MB analysis_panel을 한 번만 읽어 그림용 파생 집계를 cache/에 저장.
#  ① sensor_lu.csv (센서 토지이용)  ② dong_base_leq.csv (동 기준 Lday)
#  ③ weekly_lu.csv (토지이용별 주간 이동량·상대소음)  ④ did_weekly_ci.csv (DiD 주별 + 동클러스터 SE)
#  ⑤ binned_deciles.csv (양방향 demean 데실 + 동클러스터 SE)
import os
import numpy as np
import pandas as pd
import figstyle_v2 as FS

pan = pd.read_csv(os.path.join(FS.P, "analysis_panel.csv"), dtype={"serial": str, "dose_key": str},
                  usecols=["serial", "dose_key", "date", "Leq_day", "dLeq_day",
                           "lp_day", "lp_night", "lp_day_rel", "lp_day_logrel"])
pan["date"] = pd.to_datetime(pan["date"])
pan["week"] = pan["date"].dt.to_period("W").dt.start_time
print(f"panel {len(pan):,} rows · sensors {pan.serial.nunique()} · dongs {pan.dose_key.nunique()}")

# ---- 토지이용 분류 (post-lift 주간/야간 인구비 3분위 — phase3와 동일 정의) ----
post = pan[(pan["date"] >= "2022-07-01") & (pan["date"] <= "2023-12-31")]
comm = post.groupby("dose_key").apply(lambda g: g["lp_day"].mean() / g["lp_night"].mean(),
                                      include_groups=False).rename("comm_idx")
q1, q2 = comm.quantile([1/3, 2/3])
lu_map = comm.map(lambda c: "commercial" if c >= q2 else ("residential" if c <= q1 else "mixed"))
pan["lu"] = pan["dose_key"].map(lu_map)

sd = pd.read_csv(os.path.join(FS.P, "sensor_dong_map.csv"), dtype={"serial": str, "adm_cd": str})
sd = sd[sd["matched"] == 1].copy()
sd["lu"] = sd["adm_cd"].map(lu_map).fillna("mixed")
sd[["serial", "lat", "lon", "adm_cd", "lu"]].to_csv(os.path.join(FS.CACHE, "sensor_lu.csv"),
                                                    index=False, encoding="utf-8-sig")
print("① sensor_lu:", sd.lu.value_counts().to_dict())

# ---- 동 기준 소음 (post-lift 평균 Lday) ----
base_leq = post.groupby("dose_key")["Leq_day"].mean().rename("base_Lday")
base_leq.to_csv(os.path.join(FS.CACHE, "dong_base_leq.csv"), encoding="utf-8-sig")
print(f"② dong_base_leq: {len(base_leq)} dongs, {base_leq.min():.1f}~{base_leq.max():.1f} dB")

# ---- 토지이용별 주간 궤적 (동 동일가중 — phase3_trajectory와 동일) ----
city_wk = pan.groupby("week")["dLeq_day"].mean().rename("city_dLeq")
dw = pan.groupby(["lu", "dose_key", "week"]).agg(mob=("lp_day_rel", "mean"),
                                                 dLeq=("dLeq_day", "mean")).reset_index()
wk = dw.groupby(["lu", "week"]).agg(mob=("mob", "mean"), dLeq=("dLeq", "mean")).reset_index()
wk = wk.merge(city_wk, on="week")
wk["rel_noise"] = wk["dLeq"] - wk["city_dLeq"]
wk = wk.sort_values(["lu", "week"])
wk["rel_noise_s"] = wk.groupby("lu")["rel_noise"].transform(
    lambda s: s.rolling(4, min_periods=1, center=True).mean())
wk.to_csv(os.path.join(FS.CACHE, "weekly_lu.csv"), index=False, encoding="utf-8-sig")
print(f"③ weekly_lu: {len(wk)} rows")

# ---- DiD 주별 + 동클러스터 SE (그룹정의 = phase2b: Lv4기 lp_day_rel 3분위) ----
imp = pan[(pan["date"] >= "2021-07-12") & (pan["date"] <= "2021-09-30")].groupby("dose_key")["lp_day_rel"].mean()
t1, t3 = imp.quantile([1/3, 2/3])
grp = pd.Series(np.where(imp <= t1, "high", np.where(imp >= t3, "low", "mid")), index=imp.index)
pan["grp"] = pan["dose_key"].map(grp)
sub = pan[pan["grp"].isin(["high", "low"])]

def cluster_se(df, val):
    # 주×그룹 평균의 동클러스터 SE: sqrt( Σ_g (Σ_i∈g (y_i − ȳ))² ) / N
    ybar = df[val].mean(); N = len(df)
    s = df.assign(dev=df[val] - ybar).groupby("dose_key")["dev"].sum()
    return np.sqrt((s ** 2).sum()) / N

rows = []
for (week, g), d in sub.groupby(["week", "grp"]):
    rows.append({"week": week, "grp": g, "dLeq": d["dLeq_day"].mean(), "mob": d["lp_day_rel"].mean(),
                 "n": len(d), "se_dLeq": cluster_se(d, "dLeq_day"), "se_mob": cluster_se(d, "lp_day_rel")})
w = pd.DataFrame(rows)
piv = w.pivot(index="week", columns="grp", values=["dLeq", "mob", "n", "se_dLeq", "se_mob"])
piv = piv[(piv[("n", "high")] >= 50) & (piv[("n", "low")] >= 50)]
out = pd.DataFrame({
    "week": piv.index,
    "dLeq_diff": piv[("dLeq", "high")] - piv[("dLeq", "low")],
    "dLeq_diff_se": np.sqrt(piv[("se_dLeq", "high")] ** 2 + piv[("se_dLeq", "low")] ** 2),
    "mob_diff": piv[("mob", "high")] - piv[("mob", "low")],
    "mob_diff_se": np.sqrt(piv[("se_mob", "high")] ** 2 + piv[("se_mob", "low")] ** 2),
}).reset_index(drop=True)
out.to_csv(os.path.join(FS.CACHE, "did_weekly_ci.csv"), index=False, encoding="utf-8-sig")
# 구본과 point estimate 대조
old = pd.read_csv(os.path.join(FS.P, "phase2b_did_weekly.csv"), parse_dates=["week"])
m = out.merge(old[["week", "dLeq_diff"]], on="week", suffixes=("", "_old"))
print(f"④ did_weekly_ci: {len(out)} weeks · 구본 대비 max|Δ|={np.abs(m.dLeq_diff - m.dLeq_diff_old).max():.2e}")
r = out["mob_diff"].corr(out["dLeq_diff"]); rs = out["mob_diff"].corr(out["dLeq_diff"], method="spearman")
print(f"   주별 상관 Pearson {r:+.3f} Spearman {rs:+.3f} (본문 +0.44/+0.46 대조)")

# ---- 양방향 demean 데실 + 동클러스터 SE (phase4와 동일 demean) ----
d = pan.dropna(subset=["Leq_day", "lp_day_logrel"]).reset_index(drop=True)
g1 = pd.factorize(d["serial"].values)[0]; n1 = g1.max() + 1; c1 = np.bincount(g1, minlength=n1)
g2 = pd.factorize(d["date"].values)[0]; n2 = g2.max() + 1; c2 = np.bincount(g2, minlength=n2)

def demean(v, iters=15):
    v = v.astype(float).copy()
    for _ in range(iters):
        v -= (np.bincount(g1, v, n1) / c1)[g1]
        v -= (np.bincount(g2, v, n2) / c2)[g2]
    return v

yt = demean(d["Leq_day"].values)
xt = demean(d["lp_day_logrel"].values)
beta = np.cov(yt, xt)[0, 1] / np.var(xt)
dec = pd.qcut(xt, 10, labels=False, duplicates="drop")
bd = pd.DataFrame({"dec": dec, "xt": xt, "yt": yt, "dong": d["dose_key"].values})
rows = []
for k, dd in bd.groupby("dec"):
    ybar = dd["yt"].mean(); N = len(dd)
    s = dd.assign(dev=dd["yt"] - ybar).groupby("dong")["dev"].sum()
    rows.append({"dec": k, "xt": dd["xt"].mean(), "yt": ybar,
                 "se": np.sqrt((s ** 2).sum()) / N, "n": N})
pd.DataFrame(rows).to_csv(os.path.join(FS.CACHE, "binned_deciles.csv"), index=False, encoding="utf-8-sig")
print(f"⑤ binned_deciles: 10 deciles · 재현 β={beta:+.4f} (정본 +0.6484 대조)")
print("완료 →", FS.CACHE)
