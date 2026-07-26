# -*- coding: utf-8 -*-
# Crossref 2차 정밀 쿼리(오매칭 보정 + 추가 실제 논문). 출력 refs_raw2.json.
import json, time, urllib.request, urllib.parse
OUT = r"C:\Users\wh850\AppData\Local\Temp\refs_raw2.json"
MAILTO = "research@example.org"
QUERIES = [
 ("covid","Categorization urban acoustic environment COVID-19 lockdown Barcelona Bonet-Sola"),
 ("covid","observation impact CoViD-19 recommendation measures urban noise levels central Stockholm Rumpler"),
 ("covid","Quieted city soundscapes COVID-19 Montreal Steele Guastavino"),
 ("covid","sound of silence Granada COVID-19 lockdown Vida Manzano"),
 ("covid","effects COVID-19 lockdown noise climate Florence Italy Bartalucci"),
 ("covid","Perception acoustic environment COVID-19 lockdown Buenos Aires Maggi"),
 ("covid","Psychological wellbeing soundscape COVID-19 lockdown Erfanian Kang"),
 ("covid","noise pollution decline COVID-19 lockdown environmental review"),
 ("health","Burden disease environmental noise quantification healthy life years lost Europe"),
 ("health","Environmental noise pollution United States developing effective policy Hammer Swinburn Neitzel"),
 ("health","WHO systematic review road traffic noise annoyance Guski Schreckenberg"),
 ("health","exposure-response relationship transportation noise annoyance Miedema Oudshoorn"),
 ("sscape","ISO 12913-1 acoustics soundscape definitional framework"),
 ("sscape","Acoustic comfort evaluation urban open public spaces Yang Kang"),
 ("sscape","soundscape approach urban environmental noise management Brown"),
 ("sscape","effect audiovisual interaction soundscape Jo Jeon"),
 ("sscape","Jeon soundscape preservation quality urban park Korea"),
 ("mobility","Dynamic population mapping using mobile phone data Deville"),
 ("mobility","Seoul de facto living population estimation mobile signaling big data"),
 ("mobility","Varieties mobility measures survey mobile phone data validity"),
 ("mobility","Urban population dynamics COVID-19 pandemic mobile network Romanillos"),
 ("iot","review wireless acoustic sensor networks environmental noise monitoring Alias Alsina-Pages"),
 ("iot","smartphone crowd-sourced database environmental noise assessment Picaut"),
 ("iot","DYNAMAP dynamic acoustic mapping low-cost wireless sensor network Sevillano"),
 ("iot","Calibration low-cost sensor network environmental noise pollution monitoring"),
 ("traffic","Dynamic approach traffic noise modelling road mobility Can"),
 ("traffic","Land use regression modelling road traffic noise Morley"),
 ("method","natural experiments evaluate population health interventions Craig"),
 ("method","mostly harmless econometrics empiricist companion"),
 ("city","COVID-19 lockdown nitrogen dioxide air quality satellite Venter"),
 ("city","Temporary reduction daily global CO2 emissions COVID-19 confinement Le Quere"),
 ("city","Smart sustainable city acoustic sensor network noise WASN Zambon"),
 ("city","urban soundscape mapping GIS spatial noise Seoul"),
 ("covid","COVID-19 lockdown impact urban soundscape India Mumbai Delhi"),
]
def query(q):
    url="https://api.crossref.org/works?"+urllib.parse.urlencode(
        {"query.bibliographic":q,"rows":2,"mailto":MAILTO,
         "select":"title,author,container-title,published,volume,page,DOI,type,issued"})
    req=urllib.request.Request(url,headers={"User-Agent":f"refs/1.0 (mailto:{MAILTO})"})
    with urllib.request.urlopen(req,timeout=30) as r: return json.load(r)["message"]["items"]
def fmt(it):
    title=(it.get("title") or ["?"])[0]; au=it.get("author") or []
    a0=(au[0].get("family","?") if au else "?")+(" et al." if len(au)>1 else "")
    yr=""
    for k in ("published","issued"):
        if it.get(k,{}).get("date-parts"): yr=it[k]["date-parts"][0][0]; break
    return {"title":title,"first_author":a0,"n_authors":len(au),"year":yr,
            "journal":(it.get("container-title") or [""])[0],"volume":it.get("volume",""),
            "page":it.get("page",""),"doi":it.get("DOI",""),"type":it.get("type",""),
            "authors_full":[f"{a.get('family','')} {a.get('given','')[:1] if a.get('given') else ''}".strip() for a in au[:8]]}
results=[]
for topic,q in QUERIES:
    try:
        items=query(q); top=fmt(items[0]) if items else {"title":"(none)"}
        top["topic"]=topic; top["query"]=q[:42]; results.append(top)
        print(f"[{topic:8s}] {top['year']!s:4s} {top['first_author'][:20]:20s} | {top['journal'][:32]:32s} | {top['title'][:58]}")
    except Exception as e:
        print(f"[{topic:8s}] ERR {type(e).__name__}")
    time.sleep(0.4)
json.dump(results,open(OUT,"w",encoding="utf-8"),ensure_ascii=False,indent=1)
print(f"\n수집 {len(results)}건 -> {OUT}")
