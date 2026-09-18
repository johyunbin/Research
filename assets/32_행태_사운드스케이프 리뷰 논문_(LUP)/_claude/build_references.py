# -*- coding: utf-8 -*-
"""
Paper32 — 참고문헌 생성
① 본문에서 저자·연도로 인용한 논문의 정식 서지를 Crossref에서 확보
② 포함 98편 전체 목록(Supplementary S5용)도 별도 생성
출력: fulltext/references_intext.md · references_all_included.md · references.csv
"""
import sys, os, csv, re, json, time, unicodedata
import urllib.parse, urllib.request, urllib.error

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
UA = {"User-Agent": "paper32-SR/0.1 (mailto:wh8502@naver.com)"}

# 본문에서 저자·연도로 직접 인용한 코퍼스 논문 → uid
INTEXT = {
    "14": "Moser 1988 — roadworks noise and helping behaviour",
    "CT0025": "Mathews & Canon 1975 — lawnmower noise and helping behaviour",
    "931": "Chen 2023 — natural vs noise, group interaction",
    "1069": "Chen 2024 — natural vs noise, paired interaction",
    "461": "Franěk 2018 — birdsong vs traffic, walking speed",
    "532": "Franěk 2019 — birdsong vs city noise, walking speed",
    "617": "Oases street — nature vs traffic, field observation",
    "941": "environmental manipulation, walking speed 1.24→1.18 m/s",
    "323": "Aletta 2016 — music and dwell time",
    "665": "Ba & Kang 2020 — music and dwell time",
    "1280": "Fu 2026 — natural sound index and long stay",
    "CT0126": "Montes González 2022 — LAeq and speech disruption",
    "CT0184": "Cao & Kang 2021 — companionship and sound noticing",
    "1076": "Xu 2024 — pleasantness and static behaviour",
    "980": "Bao 2023 — dwell time and restorativeness",
    "1177": "Béjaïa — sound and walking comfort",
    "1018": "sitting/walking groups correlation",
    "1221": "Wang 2025 — natural events and queuing time",
    "CT0090": "Franěk 2014 — music tempo and walking speed",
    "CT0166": "Franěk & Režný 2021 — perceived soundscape and walking speed",
    "CT0335": "Schrapel 2022 — augmented footsteps and gait",
    # "1226" 2026-09-18 R1 재판정으로 코퍼스 배제 → 본문 인용 제거(fix_1226_citation_20260918.py)
    "951": "running intensity and traffic noise",
    "566": "Musikiosk — user-controlled music installation",
    "1272": "vehicle warning sound and crossing",
    "CT0175": "Dublin 2018 — park sound level and visit frequency",
    "CT0371": "Routhier 2024 — audible pedestrian signals, GPS trajectories",
}


# Crossref·OpenAlex 어디에도 저자가 없는 레코드 — 원문 표지에서 직접 옮긴다.
MANUAL = {
    "941": "Ba, M., Li, Z., & Kang, J. (2024). Research on the combined effects of plant odor "
           "and traffic noise on crowd behaviors in urban environments. "
           "*Landscape Architecture Frontiers*, *12*(6), 46-63. "
           "https://doi.org/10.15302/j-laf-1-020106",
}


def get(url, tries=4):
    req = urllib.request.Request(url, headers=UA)
    for i in range(tries):
        try:
            with urllib.request.urlopen(req, timeout=45) as r:
                return json.loads(r.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            if e.code == 404:
                return None
            if e.code in (429, 500, 502, 503) and i < tries - 1:
                time.sleep(3 * (i + 1)); continue
            return None
        except Exception:
            if i < tries - 1:
                time.sleep(3 * (i + 1)); continue
            return None


MOJI = re.compile(r"[ÃÄÅÂÐ×Þ][^\x00-\x7f]")


def demoji(s):
    if not s or not MOJI.search(s):
        return s
    for enc in ("cp1252", "latin-1"):
        try:
            f = s.encode(enc).decode("utf-8")
        except (UnicodeEncodeError, UnicodeDecodeError):
            continue
        if not MOJI.search(f) and "�" not in f:
            return f
    return s


def apa(d):
    """Crossref 레코드 → APA 7 저널 논문 서식(LUP는 APA 계열)"""
    auths = d.get("author") or []
    names = []
    for a in auths:
        fam = demoji(a.get("family") or a.get("name") or "")
        giv = demoji(a.get("given") or "")
        ini = " ".join(f"{p[0]}." for p in re.split(r"[\s\-]+", giv) if p)
        names.append(f"{fam}, {ini}".strip().rstrip(","))
    if len(names) > 20:
        who = ", ".join(names[:19]) + ", … " + names[-1]
    elif len(names) > 1:
        who = ", ".join(names[:-1]) + ", & " + names[-1]
    else:
        who = names[0] if names else "[Author unknown]"
    yr = ""
    for k in ("published-print", "published-online", "issued", "created"):
        p = (d.get(k) or {}).get("date-parts") or []
        if p and p[0] and p[0][0]:
            yr = str(p[0][0]); break
    title = demoji((d.get("title") or [""])[0]).rstrip(".")
    jour = demoji((d.get("container-title") or [""])[0])
    vol = d.get("volume", "")
    iss = d.get("issue", "")
    pg = d.get("page", "")
    doi = (d.get("DOI") or "").lower()
    s = f"{who} ({yr}). {title}. *{jour}*"
    if vol:
        s += f", *{vol}*"
        if iss:
            s += f"({iss})"
    if pg:
        s += f", {pg}"
    s += "."
    if doi:
        s += f" https://doi.org/{doi}"
    return s, yr, (names[0].split(",")[0] if names else "zz")


def main():
    ext = {r["uid"]: r for r in csv.DictReader(open(os.path.join(FT, "corpus_v4_extraction.csv"),
                                                    encoding="utf-8-sig"))}
    verd = {r["uid"]: r["final_verdict"] for r in
            csv.DictReader(open(os.path.join(FT, "corpus_v4_verdicts.csv"), encoding="utf-8-sig"))}
    # uid → DOI 는 여러 소스에 흩어져 있다
    doi = {}
    for p, kc, dc in [(os.path.join(FT, "ct_screen_final.csv"), "rec", "doi"),
                      (os.path.join(FT, "oa_supp_screen_final.csv"), "sid", "doi"),
                      (os.path.join(FT, "fulltext_status_20260803.csv"), "no", "doi"),
                      (os.path.join(FT, "packet_index.csv"), "no", "doi")]:
        if not os.path.exists(p):
            continue
        pre = "CT" if "ct_screen" in p else ("OAS" if "oa_supp" in p else "")
        for r in csv.DictReader(open(p, encoding="utf-8-sig")):
            k = str(r.get(kc, "")).strip()
            if not k or not r.get(dc):
                continue
            uid = f"{pre}{int(k):04d}" if pre else k
            doi.setdefault(uid, r[dc].strip())

    rows, misses = [], []
    for uid, why in INTEXT.items():
        d = doi.get(uid, "")
        if uid in MANUAL:
            m = re.match(r"^(\S+),.*?\((\d{4})\)", MANUAL[uid])
            rows.append({"uid": uid, "why": why, "doi": d, "year": m.group(2),
                         "sort": m.group(1).lower().rstrip(","), "apa": MANUAL[uid]})
            continue
        rec = get("https://api.crossref.org/works/" + urllib.parse.quote(d)) if d else None
        if rec:
            s, yr, first = apa(rec["message"])
        else:
            e = ext.get(uid, {})
            s = (f"[서지 미확보] {demoji(e.get('title', '') or '')}. *{e.get('journal', '')}*, "
                 f"{e.get('year', '')}." + (f" https://doi.org/{d}" if d else ""))
            yr, first = e.get("year", ""), "zz"
            misses.append(uid)
        rows.append({"uid": uid, "why": why, "doi": d, "year": yr,
                     "sort": first.lower(), "apa": s})
        time.sleep(0.4)

    rows.sort(key=lambda r: (r["sort"], r["year"]))
    with open(os.path.join(FT, "references.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["uid", "why", "doi", "year", "apa"],
                           extrasaction="ignore")
        w.writeheader(); w.writerows(rows)

    L = ["# Paper32 — 본문 인용 참고문헌 (APA 7)\n",
         f"\n원고 본문에서 저자·연도로 직접 인용한 코퍼스 논문 {len(rows)}편. "
         "Crossref에서 정식 서지를 받아 생성했다(재현 = `build_references.py`).\n",
         "\n⚠️ 리뷰 배경 문헌(ISO 12913, Aletta 프레임워크, Zhang 2025 등 코퍼스 밖 인용)은 "
         "여기 없다. 집필 시 별도 추가해야 한다.\n\n"]
    for r in rows:
        L.append(f"- {r['apa']}\n")
    if misses:
        L.append(f"\n## ⚠️ 서지 미확보 {len(misses)}건\n\n" +
                 "".join(f"- {m} (DOI: {doi.get(m, '없음')})\n" for m in misses))
    open(os.path.join(FT, "references_intext.md"), "w", encoding="utf-8").write("".join(L))

    # ── 포함 98편 전체 목록(Supplementary) ──
    inc = [u for u, v in verd.items() if v == "FINAL_INCLUDE"]
    L2 = [f"# Supplementary — 포함 연구 {len(inc)}편 목록\n\n",
          "| # | uid | 갈래 | 연도 | 저널 | 제목 |\n|---|---|---|---|---|---|\n"]
    SRC = {"db-search": "DB", "citation-tracking": "CT", "openalex-supplementary": "OAS"}
    for i, u in enumerate(sorted(inc, key=lambda x: (ext.get(x, {}).get("year", ""), x)), 1):
        e = ext.get(u, {})
        L2.append(f"| {i} | {u} | {SRC.get(e.get('source', ''), '?')} | {e.get('year', '')} | "
                  f"{(e.get('journal') or '')[:38]} | {demoji(e.get('title', '') or '')[:76]} |\n")
    open(os.path.join(FT, "references_all_included.md"), "w", encoding="utf-8").write("".join(L2))

    print(f"[완료] 본문 인용 {len(rows)}편 · 서지 미확보 {len(misses)}건 {misses}")
    print(f"       포함 전체 {len(inc)}편 목록 생성")
    print("[저장] references_intext.md · references_all_included.md · references.csv")


if __name__ == "__main__":
    main()
