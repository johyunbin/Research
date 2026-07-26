# -*- coding: utf-8 -*-
# refs_raw(1,2).json 통합·선별·Vancouver 포맷 -> refs_final.json (manuscript 빌더가 사용).
# INCLUDE = 검토로 선별한 실제·관련 논문(제목 부분문자열). 표준문헌은 수기 추가(WHO/ISO/Schafer).
import json, os
T = r"C:\Users\wh850\AppData\Local\Temp"
raw = json.load(open(os.path.join(T, "refs_raw.json"), encoding="utf-8")) + \
      json.load(open(os.path.join(T, "refs_raw2.json"), encoding="utf-8"))
# dedupe by doi
by = {}
for r in raw:
    k = r.get("doi") or r.get("title", "")[:40]
    if k and k not in by:
        by[k] = r

def find(sub):
    for r in by.values():
        if sub.lower() in (r.get("title") or "").lower():
            return r
    return None

def vanc(r):
    au = r.get("authors_full") or []
    au = [a for a in au if a.strip()]
    if len(au) > 6:
        astr = ", ".join(au[:6]) + ", et al"
    else:
        astr = ", ".join(au)
    t = (r.get("title") or "").rstrip(". ")
    j = r.get("journal", "")
    y = r.get("year", "")
    vp = r.get("volume", "")
    pg = r.get("page", "")
    tail = f" {y};{vp}:{pg}".rstrip(":") if y else ""
    doi = f" doi:{r['doi']}" if r.get("doi") else ""
    return f"{astr}. {t}. {j}.{tail}.{doi}".replace("..", ".")

# (key, 제목 부분문자열) — 주제/순서. 표준문헌은 manual.
INCLUDE = [
 ("who_guide", None, "World Health Organization. Environmental Noise Guidelines for the European Region. Copenhagen: WHO Regional Office for Europe; 2018."),
 ("who_burden", None, "World Health Organization. Burden of disease from environmental noise: quantification of healthy life years lost in Europe. Copenhagen: WHO Regional Office for Europe; 2011."),
 ("basner", "Auditory and non-auditory effects of noise"),
 ("munzel", "Cardiovascular effects of environmental noise"),
 ("hammer", "Environmental Noise Pollution in the United States"),
 ("guski", "Systematic Review on Environmental Noise and Annoyance"),
 ("miedema", "Annoyance from Transportation Noise"),
 ("vankamp", "Sleep-disturbance and quality of sleep in Hong Kong"),
 ("schafer", None, "Schafer RM. The Soundscape: Our Sonic Environment and the Tuning of the World. Rochester (VT): Destiny Books; 1994."),
 ("iso", None, "International Organization for Standardization. ISO 12913-1:2014 Acoustics - Soundscape - Part 1: Definition and conceptual framework. Geneva: ISO; 2014."),
 ("aletta16", "Soundscape descriptors and a conceptual framework"),
 ("axelsson", "A principal components model of soundscape perception"),
 ("kang16", "Ten questions on the soundscapes of the built environment"),
 ("yangkang", "Acoustic comfort evaluation in urban open public spaces"),
 ("kang19", "Noise Management"),
 ("hong17", "Relationship between spatiotemporal variability of soundsc"),
 ("erfanian", "Psychological Well-being and Demographic Factors"),
 ("basu", "Investigating changes in noise pollution"),
 ("asensio", "Changes in noise levels in the city of Madrid"),
 ("aletta20", "Assessing the changing urban sound environment"),
 ("rumpler", "observation of the impact of CoViD"),
 ("steele", "Quieted City Sound"),
 ("manzano", "sound of silence"),
 ("maggi", "Perception of the acoustic environment during COVID"),
 ("montano", "SOUNDSCAPE CHANGES DUE TO THE LOCKDOWN"),
 ("mishra", "lockdown on noise pollution levels"),
 ("sonaviya", "Integrated road traffic noise mapping in urban Indian"),
 ("lequere", "Temporary reduction in daily global CO"),
 ("brancher", "Increased ozone pollution alongside reduced nitrogen"),
 ("mahato", "Revisiting air quality during lockdown"),
 ("deville", "Dynamic population mapping using mobile phone data"),
 ("paez", "Using Google Community Mobility Reports"),
 ("kalleitner", "Varieties of mobility measures"),
 ("romanillos", "Urban population dynamics during the COVID"),
 ("yim", "Two-Timescale Typology of Neighborhood-Scale Commercial"),
 ("mydlarz", "implementation of low-cost urban acoustic monitoring"),
 ("bello", "SONYC"),
 ("alias", "Review of Wireless Acoustic Sensor Networks for Environmental"),
 ("sevillano", "DYNAMAP"),
 ("alsina", "Smart Wireless Acoustic Sensor Network Design for Noise"),
 ("boumchich", "Clustering Method to Detect Spatial Events"),
 ("peng", "Hierarchical Wireless Acoustic Sensor Network"),
 ("cui", "calibration system for low-cost Sensor Network in air"),
 ("adulaimi", "Traffic Noise Modelling Using Land Use Regression"),
 ("gharehchahi", "Geospatial analysis for environmental noise mapping"),
 ("vogiatzis", "Soundscape design guidelines through noise mapping"),
 ("abadie", "When Should You Adjust Standard Errors for Clustering"),
 ("craig", "Using natural experiments to evaluate population health"),
 ("guo", "Fixed effects spatial panel data models"),
 ("bartalucci", "survey on the soundscape perception before and during"),
 ("picaut", "Exploiting data from the NoiseCapture application"),
 ("torresin", "Indoor soundscape assessment: A principal components"),
]

final = []
miss = []
for entry in INCLUDE:
    key = entry[0]
    if len(entry) == 3 and entry[1] is None:  # manual standard ref
        final.append({"key": key, "vancouver": entry[2], "manual": True})
        continue
    sub = entry[1]
    r = find(sub)
    if r and r.get("authors_full"):
        final.append({"key": key, "vancouver": vanc(r), "doi": r.get("doi", ""), "year": r.get("year"), "journal": r.get("journal")})
    else:
        miss.append((key, sub))

for i, f in enumerate(final, 1):
    f["num"] = i
json.dump(final, open(os.path.join(T, "refs_final.json"), "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(f"최종 참고문헌 {len(final)}개")
for f in final:
    print(f"  [{f['num']:2d}] {f['key']:12s} {f['vancouver'][:95]}")
if miss:
    print("\n미발견(보강 필요):", miss)
