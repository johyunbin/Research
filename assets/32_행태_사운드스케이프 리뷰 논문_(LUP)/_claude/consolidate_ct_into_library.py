# -*- coding: utf-8 -*-
"""
Paper32 — 인용추적 PDF를 수집논문_PDF 라이브러리에 같은 명명 규칙으로 편입
명명: 년도_저널축약_저자(Jo et al 형식)_제목.pdf  (match_and_rename_include.py와 동일 규칙)
⚠️ 인용추적분은 아직 전문심사 전 '후보'다. 파일은 한 폴더에 모으되 출처·상태를
   fulltext/pdf_library_manifest.csv 에 기록해 본코퍼스 100편과 구분 가능하게 둔다.
"""
import sys, os, csv, re, json, time, shutil, unicodedata
import urllib.request, urllib.parse, urllib.error

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
SRC = os.path.join(BASE, "ct_pdf")
LIB = os.path.join(os.path.dirname(BASE), "수집논문_PDF")
UA = {"User-Agent": "paper32-SR/0.1 (mailto:wh8502@naver.com)"}

ABBR = [
    ("SCIENCE OF THE TOTAL ENVIRONMENT", "STOTEN"), ("BUILDING AND ENVIRONMENT", "BuildEnv"),
    ("LANDSCAPE AND URBAN PLANNING", "LUP"), ("APPLIED ACOUSTICS", "ApplAcoust"),
    ("JOURNAL OF THE ACOUSTICAL SOCIETY OF AMERICA", "JASA"),
    ("URBAN FORESTRY & URBAN GREENING", "UFUG"), ("URBAN FORESTRY AND URBAN GREENING", "UFUG"),
    ("JOURNAL OF TRANSPORT & HEALTH", "JTH"), ("JOURNAL OF TRANSPORT AND HEALTH", "JTH"),
    ("JOURNAL OF TRANSPORT GEOGRAPHY", "JTransGeogr"), ("TRANSPORT POLICY", "TransPolicy"),
    ("APPLIED GEOGRAPHY", "ApplGeogr"), ("ECOLOGICAL INFORMATICS", "EcolInform"),
    ("ECOLOGICAL INDICATORS", "EcolIndic"), ("SUSTAINABLE CITIES AND SOCIETY", "SCS"),
    ("JOURNAL OF ENVIRONMENTAL PSYCHOLOGY", "JEnvPsych"), ("ENVIRONMENT INTERNATIONAL", "EnvInt"),
    ("INTERNATIONAL JOURNAL OF ENVIRONMENTAL RESEARCH AND PUBLIC HEALTH", "IJERPH"),
    ("TRANSPORTATION RESEARCH", "TransRes"), ("ENVIRONMENTAL POLLUTION", "EnvPollut"),
    ("HEALTH & PLACE", "HealthPlace"), ("CITIES", "Cities"), ("LANDSCAPE ECOLOGY", "LandEcol"),
    ("ENVIRONMENT AND BEHAVIOR", "EnvBehav"), ("FRONTIERS IN PSYCHOLOGY", "FrontPsych"),
    ("SCIENTIFIC REPORTS", "SciRep"), ("SUSTAINABILITY", "Sustainability"),
    ("APPLIED SCIENCES", "ApplSci"), ("FORESTS", "Forests"), ("ACOUSTICS", "Acoustics"),
    ("NOISE MAPPING", "NoiseMapp"), ("PLOS ONE", "PLOSONE"),
    ("JOURNAL OF URBAN DESIGN", "JUrbanDes"), ("JOURNAL OF URBAN MANAGEMENT", "JUrbanMgmt"),
    ("JOURNAL OF PERSONALITY AND SOCIAL PSYCHOLOGY", "JPSP"),
    ("JOURNAL OF OUTDOOR RECREATION AND TOURISM", "JORT"),
    ("JOURNAL OF RESOURCES AND ECOLOGY", "JResEcol"),
    ("TRAVEL BEHAVIOUR AND SOCIETY", "TravBehavSoc"),
    ("TOURISM RECREATION RESEARCH", "TourRecRes"),
    ("ISPRS INTERNATIONAL JOURNAL OF GEO-INFORMATION", "IJGI"),
    ("INFRASTRUCTURES", "Infrastructures"), ("SENSORS", "Sensors"), ("LAND", "Land"),
]


def abbrev(journal):
    j = (journal or "").upper()
    for full, ab in ABBR:
        if full in j:
            return ab
    words = re.findall(r"[A-Za-z]+", journal or "")
    return "".join(w[:4].capitalize() for w in words[:2]) or "Journal"


# UTF-8을 CP1252/Latin-1로 잘못 읽은 흔적: 선행 바이트가 Ã Ä Å Â Ð × Þ 이고 뒤에 비ASCII가 붙는다.
MOJI = re.compile(r"[ÃÄÅÂÐ×Þ][^\x00-\x7f]")


def demojibake(s):
    """Crossref 원본에 mojibake가 그대로 저장된 경우가 있다(예: 'FranÄ›k' ← 'Franěk').
    UTF-8 바이트가 CP1252로 읽힌 것이 대부분이라 cp1252 → latin-1 순으로 왕복 복구를 시도한다.
    복구 결과에 흔적이 남거나 디코딩이 실패하면 원문을 그대로 둔다(과교정 방지)."""
    if not s or not MOJI.search(s):
        return s
    for enc in ("cp1252", "latin-1"):
        try:
            fixed = s.encode(enc).decode("utf-8")
        except (UnicodeEncodeError, UnicodeDecodeError):
            continue
        if not MOJI.search(fixed) and "�" not in fixed:
            return fixed
    return s


def safe(s, n=70):
    s = demojibake(s or "")
    s = re.sub(r'[\\/:*?"<>|]', "", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s[:n].rstrip(" .")


def crossref(doi):
    try:
        req = urllib.request.Request(
            f"https://api.crossref.org/works/{urllib.parse.quote(doi)}", headers=UA)
        d = json.loads(urllib.request.urlopen(req, timeout=30).read().decode("utf-8"))["message"]
    except Exception:
        return None
    auths = d.get("author") or []
    first = (auths[0].get("family") or auths[0].get("name") or "") if auths else ""
    n = len(auths)
    if n == 0:
        astr = "Anon"
    elif n == 1:
        astr = first
    elif n == 2:
        second = auths[1].get("family") or auths[1].get("name") or ""
        astr = f"{first} & {second}"
    else:
        astr = f"{first} et al"
    yr = ""
    for k in ("published-print", "published-online", "issued", "created"):
        p = (d.get(k) or {}).get("date-parts") or []
        if p and p[0] and p[0][0]:
            yr = str(p[0][0]); break
    return {"year": yr, "journal": (d.get("container-title") or [""])[0],
            "authors": astr, "title": (d.get("title") or [""])[0]}


def main():
    os.makedirs(LIB, exist_ok=True)
    fin = {r["rec"]: r for r in csv.DictReader(open(os.path.join(FT, "ct_screen_final.csv"),
                                                    encoding="utf-8-sig"))}
    files = sorted(f for f in os.listdir(SRC) if f.lower().endswith(".pdf"))
    print(f"[0] ct_pdf {len(files)}편 · 라이브러리 현재 {len(os.listdir(LIB))}개")

    rows, renamed, failed = [], 0, []
    for f in files:
        m = re.match(r"CT(\d{4})_", f)
        if not m:
            failed.append((f, "REC 파싱 실패")); continue
        rec = str(int(m.group(1)))
        r = fin.get(rec)
        if not r:
            failed.append((f, "ct_screen_final에 없음")); continue
        meta = crossref(r["doi"]) if r["doi"] else None
        time.sleep(0.35)
        if not meta or not meta["title"]:
            meta = {"year": r.get("year", ""), "journal": r.get("journal", ""),
                    "authors": "Anon", "title": r["title"]}
            src_meta = "fallback"
        else:
            src_meta = "crossref"
        newname = (f"{meta['year'] or r.get('year','')}_{abbrev(meta['journal'] or r.get('journal',''))}"
                   f"_{safe(meta['authors'], 40)}_{safe(meta['title'], 70)}.pdf")
        dst = os.path.join(LIB, newname)
        if os.path.exists(dst):
            print(f"    = 이미 있음 {newname[:66]}")
        else:
            shutil.copy2(os.path.join(SRC, f), dst)
            renamed += 1
            print(f"    ✓ REC {rec:>4} → {newname[:70]}")
        rows.append({"rec": rec, "source": "citation-tracking", "screen": r["final"],
                     "status": "candidate — 전문심사 전", "doi": r["doi"],
                     "year": meta["year"], "journal": meta["journal"],
                     "authors": meta["authors"], "meta_from": src_meta,
                     "filename": newname, "orig_ct_file": f})

    # 본코퍼스 100편도 매니페스트에 기록(구분 가능하게)
    pm = {}
    p = os.path.join(FT, "pdf_map.csv")
    if os.path.exists(p):
        for r in csv.DictReader(open(p, encoding="utf-8-sig")):
            pm[r.get("filename") or r.get("file") or ""] = r
    verd = {}
    for r in csv.DictReader(open(os.path.join(FT, "ft_verdicts_v2.csv"), encoding="utf-8-sig")):
        verd[r["no"]] = r["final_verdict"]
    ct_names = {x["filename"] for x in rows}
    for fn in sorted(os.listdir(LIB)):
        if not fn.lower().endswith(".pdf") or fn in ct_names:
            continue
        r = pm.get(fn, {})
        no = r.get("no", "")
        rows.append({"rec": no, "source": "main-search", "screen": "",
                     "status": verd.get(no, "확인 필요"), "doi": r.get("doi", ""),
                     "year": fn[:4], "journal": "", "authors": "",
                     "meta_from": "pdf_map", "filename": fn, "orig_ct_file": ""})

    with open(os.path.join(FT, "pdf_library_manifest.csv"), "w", newline="",
              encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["rec", "source", "screen", "status", "doi", "year",
                                          "journal", "authors", "meta_from", "filename",
                                          "orig_ct_file"])
        w.writeheader(); w.writerows(rows)

    from collections import Counter
    print(f"\n[완료] 신규 편입 {renamed}편 · 라이브러리 총 {len(os.listdir(LIB))}개")
    print(f"       출처: {dict(Counter(x['source'] for x in rows))}")
    if failed:
        print("       ⚠️ 처리 실패:", failed)
    print(f"       매니페스트 → fulltext/pdf_library_manifest.csv")


if __name__ == "__main__":
    main()
