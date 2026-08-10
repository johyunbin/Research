# -*- coding: utf-8 -*-
"""
Paper32 — Introduction 배경 인용 문헌의 **실재 확인**

서론에 인용을 넣으라는 요청에 따라 배경 문헌을 넣되, **지어내지 않는다**.
후보 질의를 Crossref 에 던져 실제 레코드를 받고, 제목·연도·저널이 기대와 맞는지
사람이 확인할 수 있게 후보를 나열한다. 확인된 것만 원고에 쓴다.

출력: fulltext/background_refs_candidates.md
"""
import sys, os, json, time, re
import urllib.parse, urllib.request, urllib.error

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
UA = {"User-Agent": "paper32-SR/0.1 (mailto:wh8502@naver.com)"}

# (내부 키, 검색 질의, 기대 연도, 기대 저널·출판사 힌트)
QUERIES = [
    ("soundscape_iso_framework",
     "Soundscape descriptors and a conceptual framework for developing predictive soundscape models",
     2016, "Landscape and Urban Planning"),
    ("axelsson_pca",
     "A principal components model of soundscape perception", 2010,
     "Journal of the Acoustical Society of America"),
    ("kang_ten_questions",
     "Ten questions on the soundscapes of the built environment", 2016, "Building and Environment"),
    ("iso12913_review",
     "The Soundscape Indices (SSID) protocol a method for urban soundscape surveys", 2020,
     "Applied Sciences"),
    ("basner_health",
     "Auditory and non-auditory effects of noise on health", 2014, "Lancet"),
    ("who_burden",
     "Burden of disease from environmental noise", 2011, "WHO"),
    ("prisma2020",
     "The PRISMA 2020 statement an updated guideline for reporting systematic reviews", 2021, "BMJ"),
    ("mmat2018",
     "The Mixed Methods Appraisal Tool (MMAT) version 2018 for information professionals and researchers",
     2018, "Education for Information"),
    ("hedges_g",
     "Distribution theory for Glass's estimator of effect size and related estimators", 1981,
     "Journal of Educational Statistics"),
    ("chinn",
     "A simple method for converting an odds ratio to effect size for use in meta-analysis", 2000,
     "Statistics in Medicine"),
    ("hartung_knapp",
     "On tests of the overall treatment effect in meta-analysis with normally distributed responses",
     2001, "Statistics in Medicine"),
    ("hts_pi",
     "A re-evaluation of random-effects meta-analysis", 2009,
     "Journal of the Royal Statistical Society Series A"),
    ("zhang_lup_2025",
     "landscape perception meta-analysis restorative", 2025, "Landscape and Urban Planning"),
    ("soundscape_physical_activity",
     "soundscape physical activity systematic review urban green space", 2026, "Cities & Health"),
    ("natural_sounds_health",
     "A synthesis of health benefits of natural sounds and their distribution in national parks",
     2021, "PNAS"),
    ("gehl_public_life",
     "public life street behaviour observation urban design", 0, ""),
    ("noise_annoyance_meta",
     "Annoyance from transportation noise relationships with exposure metrics", 2001,
     "Environmental Health Perspectives"),
    ("restorative_soundscape",
     "The effects of sound on perceived restorativeness in urban green space", 0, ""),
]


def crossref(q, rows=4):
    url = ("https://api.crossref.org/works?rows=%d&select=DOI,title,author,issued,"
           "container-title,type&query.bibliographic=%s" % (rows, urllib.parse.quote(q)))
    req = urllib.request.Request(url, headers=UA)
    for i in range(3):
        try:
            with urllib.request.urlopen(req, timeout=45) as r:
                return json.loads(r.read().decode("utf-8"))["message"]["items"]
        except Exception:
            if i < 2:
                time.sleep(2 * (i + 1)); continue
    return []


def fmt(it):
    yr = ""
    p = (it.get("issued") or {}).get("date-parts") or []
    if p and p[0] and p[0][0]:
        yr = str(p[0][0])
    au = it.get("author") or []
    who = (au[0].get("family", "?") if au else "?") + (" et al." if len(au) > 1 else "")
    ti = (it.get("title") or [""])[0]
    ct = (it.get("container-title") or [""])[0]
    return f"{who} ({yr}) — {ti[:96]} — *{ct[:44]}* — {it.get('DOI')}"


def main():
    L = ["# Introduction 배경 인용 후보 (Crossref 실재 확인)\n\n",
         "⚠️ **여기 실린 것은 후보일 뿐이다.** 제목·연도·저널이 의도한 문헌과 맞는지 확인한 뒤에만 "
         "원고에 쓴다. 확인되지 않은 것은 인용하지 않는다.\n\n"]
    found = {}
    for key, q, yr, hint in QUERIES:
        items = crossref(q)
        L.append(f"## {key}\n\n질의: `{q}`  (기대 {yr or '?'} · {hint or '?'})\n\n")
        if not items:
            L.append("- ❌ 결과 없음\n\n"); print(f"❌ {key}"); continue
        for it in items:
            L.append(f"- {fmt(it)}\n")
        L.append("\n")
        top = items[0]
        found[key] = top
        p = (top.get("issued") or {}).get("date-parts") or []
        y0 = p[0][0] if (p and p[0] and p[0][0]) else 0
        ok = (yr == 0) or (abs(y0 - yr) <= 1)
        print(f"{'✅' if ok else '⚠️ '} {key}: {fmt(top)[:110]}")
        time.sleep(0.6)

    open(os.path.join(FT, "background_refs_candidates.md"), "w", encoding="utf-8").write("".join(L))
    print(f"\n[저장] fulltext/background_refs_candidates.md ({len(found)}/{len(QUERIES)} 후보 확보)")


if __name__ == "__main__":
    main()
