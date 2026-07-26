# -*- coding: utf-8 -*-
# Crossref API로 후보 논문 메타데이터 검증 수집 (날조 방지). query.bibliographic 최상위 매칭 반환.
# 출력: 검토용 출력 + refs_raw.json. 관련성은 사람(Claude)이 검토 후 선별.
import json, time, urllib.request, urllib.parse, os

OUT = r"C:\Users\wh850\AppData\Local\Temp\refs_raw.json"
MAILTO = "research@example.org"

# (topic, query) — 실제 존재 가능성 높은 논문/주제. 매칭 결과를 검토해 선별.
QUERIES = [
 ("covid","Investigating changes in noise pollution due to the COVID-19 lockdown Dublin Ireland"),
 ("covid","Changes in noise levels in the city of Madrid during COVID-19 lockdown"),
 ("covid","Assessing the changing urban sound environment during the COVID-19 lockdown short-term acoustic measurements London"),
 ("covid","COVID-19 lockdown effect on urban noise Barcelona soundscape"),
 ("covid","Noise pollution variations COVID-19 lockdown Italy"),
 ("covid","Road traffic noise reduction COVID-19 lockdown Stockholm"),
 ("covid","Environmental noise pollution COVID-19 lockdown India megacity"),
 ("covid","Soundscape changes COVID-19 lockdown perception survey"),
 ("covid","Urban noise mapping COVID-19 lockdown traffic"),
 ("covid","Acoustic environment pandemic confinement measurements city"),
 ("covid","COVID-19 lockdown noise Buenos Aires Argentina"),
 ("covid","Impact of COVID-19 on environmental noise levels review"),
 ("health","Environmental Noise Guidelines for the European Region World Health Organization"),
 ("health","Burden of disease from environmental noise quantification healthy life years"),
 ("health","Auditory and non-auditory effects of noise on health Basner Lancet"),
 ("health","The adverse effects of environmental noise exposure on the cardiovascular system Munzel"),
 ("health","Transportation noise and cardiovascular disease meta-analysis"),
 ("health","Road traffic noise annoyance exposure-response relationship"),
 ("health","Night-time noise sleep disturbance exposure response"),
 ("sscape","Soundscape descriptors and a conceptual framework for developing predictive models"),
 ("sscape","A principal components model of soundscape perception"),
 ("sscape","Ten questions on the soundscapes of the built environment"),
 ("sscape","Soundscape of European cities and landscapes harmonising"),
 ("sscape","Acoustic environment and soundscape urban open public spaces"),
 ("sscape","Towards soundscape indices machine learning"),
 ("sscape","Virtual reality soundscape evaluation laboratory ecological validity"),
 ("mobility","Estimating ambient de facto population using mobile phone data"),
 ("mobility","Human mobility patterns during COVID-19 mobile phone data"),
 ("mobility","Seoul de facto population big data spatiotemporal"),
 ("mobility","Mobile phone signaling data urban population dynamics estimation"),
 ("mobility","Google COVID-19 community mobility reports validity"),
 ("mobility","Dynamic population mapping mobile network data"),
 ("iot","The implementation of low-cost urban acoustic monitoring devices"),
 ("iot","SONYC cyber-physical system monitoring mitigation urban noise pollution"),
 ("iot","Wireless acoustic sensor network environmental noise monitoring"),
 ("iot","Low-cost noise sensor calibration accuracy environmental monitoring"),
 ("iot","Internet of Things smart city environmental monitoring sustainability"),
 ("iot","Sensor drift calibration low-cost air quality sensors long-term"),
 ("iot","Urban sound monitoring distributed sensor network deep learning"),
 ("iot","Barcelona smart city noise sensor network platform"),
 ("traffic","Common noise assessment methods in Europe CNOSSOS-EU road traffic"),
 ("traffic","Relationship between traffic flow volume and road traffic noise level"),
 ("traffic","Strategic noise mapping urban road traffic agglomeration"),
 ("traffic","Urban form land use and environmental noise exposure"),
 ("traffic","Dynamic road traffic noise modelling mobility"),
 ("method","Natural experiments evaluation population health interventions"),
 ("method","Fixed effects panel data models environmental economics"),
 ("method","Difference-in-differences environmental policy evaluation"),
 ("method","When should you adjust standard errors clustering"),
 ("city","Smart sustainable cities urban data sensing framework"),
 ("city","Urban green space and noise mitigation ecosystem services"),
 ("city","Spatiotemporal variability soundscape urban morphology Seoul"),
 ("city","Land use regression environmental noise modelling city"),
 ("city","COVID-19 lockdown air quality improvement urban"),
 ("covid","Soundscape pleasantness eventfulness COVID-19 lockdown change"),
]

def query(q):
    url = "https://api.crossref.org/works?" + urllib.parse.urlencode(
        {"query.bibliographic": q, "rows": 2, "mailto": MAILTO, "select":
         "title,author,container-title,published,volume,page,DOI,type,issued"})
    req = urllib.request.Request(url, headers={"User-Agent": f"refs/1.0 (mailto:{MAILTO})"})
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.load(r)["message"]["items"]

def fmt(it):
    title = (it.get("title") or ["?"])[0]
    au = it.get("author") or []
    a0 = (au[0].get("family", "?") if au else "?") + (" et al." if len(au) > 1 else "")
    yr = ""
    for k in ("published", "issued"):
        if it.get(k, {}).get("date-parts"):
            yr = it[k]["date-parts"][0][0]; break
    jr = (it.get("container-title") or [""])[0]
    return {"title": title, "first_author": a0, "n_authors": len(au), "year": yr,
            "journal": jr, "volume": it.get("volume", ""), "page": it.get("page", ""),
            "doi": it.get("DOI", ""), "type": it.get("type", ""),
            "authors_full": [f"{a.get('family','')} {a.get('given','')[:1] if a.get('given') else ''}".strip() for a in au[:8]]}

results = []
for topic, q in QUERIES:
    try:
        items = query(q)
        top = fmt(items[0]) if items else {"title": "(none)"}
        top["topic"] = topic; top["query"] = q[:45]
        results.append(top)
        print(f"[{topic:8s}] {top['year']!s:4s} {top['first_author'][:22]:22s} | {top['journal'][:34]:34s} | {top['title'][:60]}")
    except Exception as e:
        print(f"[{topic:8s}] ERR {q[:40]}: {type(e).__name__}")
    time.sleep(0.4)

json.dump(results, open(OUT, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
print(f"\n수집 {len(results)}건 -> {OUT}")
