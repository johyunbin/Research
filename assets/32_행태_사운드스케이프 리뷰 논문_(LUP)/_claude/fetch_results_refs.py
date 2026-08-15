# -*- coding: utf-8 -*-
"""
Paper32 — Results 절 인용 후보의 **실재 확인** (`fetch_background_refs.py` 의 Results 판)

Results 에는 코퍼스 데이터로 뒷받침되지 않는 외부 사실 주장이 몇 군데 있다. 거기에 인용을
붙이되 **지어내지 않는다.** 후보 질의를 Crossref 에 던져 실제 레코드를 받고, 제목·연도·저널이
기대와 맞는지 사람이 확인할 수 있게 나열한다. 확인된 것만 원고에 쓴다.

출력: fulltext/results_refs_candidates.md
"""
import sys, os, json, time
import urllib.parse, urllib.request

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
UA = {"User-Agent": "paper32-SR/0.1 (mailto:wh8502@naver.com)"}

# (주장 키, 검색 질의, 기대 연도(0=무관), 기대 저널 힌트)
QUERIES = [
    # A. §3.1 — OpenAlex 색인 범위
    ("A_openalex_coverage",
     "OpenAlex a fully-open index of scholarly works authors venues institutions concepts", 2022, "Scientometrics"),
    ("A_openalex_vs_wos",
     "reference coverage analysis of OpenAlex compared to Web of Science and Scopus", 0, "Scientometrics"),
    ("A_db_coverage_compare",
     "multidisciplinary comparison of coverage Scopus Web of Science Dimensions citations", 2021, "Scientometrics"),
    # B. §3.1 — 제목만으로 하는 스크리닝
    ("B_title_only_screening",
     "Titles versus titles and abstracts for initial screening of articles for systematic reviews", 2013, "Clinical Epidemiology"),
    # C. §3.2 — 사운드스케이프 평가의 문화 조건화
    ("C_soundscape_crosscultural",
     "cross-cultural comparison of soundscape perception in urban open spaces", 0, ""),
    ("C_soundscape_attribute_translation",
     "soundscape attribute translation cross-cultural validation of perceptual attributes", 0, ""),
    ("C_social_demographic_sound",
     "Effects of social demographical and behavioral factors on the sound level evaluation in urban open spaces", 2008, "Journal of the Acoustical Society of America"),
    # D. §3.7 — 항공기 소음 문헌의 규모
    ("D_aircraft_annoyance_who",
     "WHO Environmental Noise Guidelines for the European Region a systematic review on environmental noise and annoyance", 2017, "IJERPH"),
    ("D_aircraft_noise_annoyance_meta",
     "aircraft noise annoyance exposure-response relationship meta-analysis", 0, ""),
    # E. §3.6 — 센싱·GPS 기반 행태 측정
    ("E_gps_behaviour_urban",
     "GPS tracking of visitor behaviour in urban green space methods review", 0, ""),
    ("E_selfreport_vs_objective",
     "agreement between self-reported and objectively measured physical activity", 0, ""),
]


def crossref(q, rows=5):
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
    return f"{who} ({yr}) — {ti[:100]} — *{ct[:46]}* — {it.get('DOI')}"


def main():
    L = ["# Results 절 인용 후보 (Crossref 실재 확인)\n\n",
         "⚠️ **여기 실린 것은 후보일 뿐이다.** 제목·연도·저널이 의도한 문헌과 맞는지 확인한 뒤에만 "
         "원고에 쓴다. 확인되지 않은 것은 인용하지 않는다.\n\n"]
    for key, q, yr, hint in QUERIES:
        items = crossref(q)
        L.append(f"## {key}\n\n질의: `{q}`  (기대 {yr or '?'} · {hint or '?'})\n\n")
        if not items:
            L.append("- ❌ 결과 없음\n\n"); print(f"❌ {key}"); continue
        for it in items:
            L.append(f"- {fmt(it)}\n")
        L.append("\n")
        top = items[0]
        p = (top.get("issued") or {}).get("date-parts") or []
        y0 = p[0][0] if (p and p[0] and p[0][0]) else 0
        ok = (yr == 0) or (abs(y0 - yr) <= 1)
        print(f"{'✅' if ok else '⚠️ '} {key}: {fmt(top)[:118]}")
        time.sleep(0.6)

    open(os.path.join(FT, "results_refs_candidates.md"), "w", encoding="utf-8").write("".join(L))
    print(f"\n[저장] fulltext/results_refs_candidates.md")


if __name__ == "__main__":
    main()
