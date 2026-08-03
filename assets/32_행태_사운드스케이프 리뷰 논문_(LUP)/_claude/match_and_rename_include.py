# -*- coding: utf-8 -*-
"""
Paper32 — Include 폴더 PDF 매칭·리네이밍
① 각 PDF에서 DOI(본문/메타) 추출 → NEEDED 목록과 DOI 매칭, 실패 시 제목 유사도 매칭
② Crossref로 저자·연도·저널 보강 (DOI 확정분)
③ 파일명 = 년도_저널축약_저자(Jo et al)_제목  로 리네임(사본을 renamed/에 생성; 원본 보존)
④ ID_XXXX.pdf 사본을 fulltext/pdf/ 에 배치(파이프라인 합류)
출력: fulltext/include_match_log.csv
"""
import sys, os, re, csv, json, glob, time, shutil, unicodedata, urllib.request, urllib.parse
import fitz

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
PROJ = os.path.dirname(BASE)
INC = os.path.join(PROJ, "Include")
FT = os.path.join(BASE, "fulltext")
OUT = os.path.join(INC, "renamed")
os.makedirs(OUT, exist_ok=True)
UA = {"User-Agent": "paper32-SR/0.1 (academic systematic review; mailto:noreply@example.org)"}

ABBR = [
    ("SCIENCE OF THE TOTAL ENVIRONMENT", "STOTEN"), ("BUILDING AND ENVIRONMENT", "BuildEnv"),
    ("LANDSCAPE AND URBAN PLANNING", "LUP"), ("APPLIED ACOUSTICS", "ApplAcoust"),
    ("JOURNAL OF THE ACOUSTICAL SOCIETY OF AMERICA", "JASA"),
    ("URBAN FORESTRY & URBAN GREENING", "UFUG"), ("URBAN FORESTRY AND URBAN GREENING", "UFUG"),
    ("JOURNAL OF TRANSPORT & HEALTH", "JTH"), ("JOURNAL OF TRANSPORT AND HEALTH", "JTH"),
    ("ECOLOGICAL INDICATORS", "EcolIndic"), ("SUSTAINABLE CITIES AND SOCIETY", "SCS"),
    ("JOURNAL OF ENVIRONMENTAL PSYCHOLOGY", "JEnvPsych"), ("ENVIRONMENT INTERNATIONAL", "EnvInt"),
    ("INTERNATIONAL JOURNAL OF ENVIRONMENTAL RESEARCH AND PUBLIC HEALTH", "IJERPH"),
    ("TRANSPORTATION RESEARCH", "TransRes"), ("ENVIRONMENTAL POLLUTION", "EnvPollut"),
    ("HEALTH & PLACE", "HealthPlace"), ("CITIES", "Cities"), ("LANDSCAPE ECOLOGY", "LandEcol"),
    ("ENVIRONMENT AND BEHAVIOR", "EnvBehav"), ("FRONTIERS IN PSYCHOLOGY", "FrontPsych"),
    ("SCIENTIFIC REPORTS", "SciRep"), ("SUSTAINABILITY", "Sustainability"),
    ("APPLIED SCIENCES", "ApplSci"), ("FORESTS", "Forests"), ("ACOUSTICS", "Acoustics"),
    ("NOISE MAPPING", "NoiseMapp"), ("PLOS ONE", "PLOSONE"),
]


def norm_title(t):
    t = unicodedata.normalize("NFKD", (t or "").lower())
    return re.sub(r"[^a-z0-9]", "", t)


def abbrev(journal):
    j = (journal or "").upper()
    for full, ab in ABBR:
        if full in j:
            return ab
    words = re.findall(r"[A-Za-z]+", journal or "")
    return "".join(w[:4].capitalize() for w in words[:2]) or "Journal"


def safe(s, n=70):
    s = re.sub(r"[\\/:*?\"<>|]", "", s or "")
    s = re.sub(r"\s+", " ", s).strip()
    return s[:n].rstrip(" .")


def extract_doi(path):
    try:
        doc = fitz.open(path)
        txt = "\n".join(doc[i].get_text() for i in range(min(3, len(doc))))
        meta = " ".join(str(v) for v in (doc.metadata or {}).values())
        doc.close()
    except Exception:
        return None, ""
    blob = txt + " " + meta
    m = re.search(r"\b(10\.\d{4,9}/[^\s\"'<>,;)\]]+)", blob)
    doi = m.group(1).rstrip(".,;)") .lower() if m else None
    return doi, txt[:1500]


def crossref(doi):
    try:
        req = urllib.request.Request(f"https://api.crossref.org/works/{urllib.parse.quote(doi)}", headers=UA)
        d = json.loads(urllib.request.urlopen(req, timeout=30).read().decode("utf-8"))["message"]
        auths = d.get("author") or []
        first = ""
        if auths:
            first = auths[0].get("family") or auths[0].get("name") or ""
        n = len(auths)
        if n == 1:
            astr = first
        elif n == 2:
            second = auths[1].get("family") or ""
            astr = f"{first} & {second}"
        else:
            astr = f"{first} et al"
        year = ""
        for k in ("published-print", "published-online", "issued", "created"):
            p = d.get(k, {}).get("date-parts", [[None]])[0][0]
            if p:
                year = str(p); break
        return {"authors": astr, "year": year,
                "journal": (d.get("container-title") or [""])[0],
                "title": (d.get("title") or [""])[0]}
    except Exception:
        return None


def main():
    needed = list(csv.DictReader(open(os.path.join(FT, "NEEDED_PDFS_priority.csv"), encoding="utf-8-sig")))
    by_doi = {(r["doi"] or "").lower(): r for r in needed if r["doi"]}
    by_title = {norm_title(r["title"]): r for r in needed}

    pdfs = sorted(glob.glob(os.path.join(INC, "*.pdf")))
    print(f"Include PDF {len(pdfs)}개 · NEEDED {len(needed)}건")

    log, matched, unmatched = [], 0, []
    for p in pdfs:
        base = os.path.basename(p)
        doi, head = extract_doi(p)
        rec, how = None, ""
        if doi and doi in by_doi:
            rec, how = by_doi[doi], "doi"
        if not rec:
            nt = norm_title(head)
            for t, r in by_title.items():
                if len(t) > 40 and t[:60] in nt:
                    rec, how = r, "title"; break
        if not rec:
            unmatched.append((base, doi or "no-doi"))
            log.append({"file": base, "matched": "NO", "how": "", "no": "", "doi": doi or "",
                        "newname": "", "note": "매칭 실패"})
            continue

        matched += 1
        cr = crossref(rec["doi"]) if rec["doi"] else None
        time.sleep(0.4)
        year = (cr or {}).get("year") or rec["year"]
        journal = (cr or {}).get("journal") or rec["journal"]
        authors = (cr or {}).get("authors") or "Unknown"
        title = (cr or {}).get("title") or rec["title"]
        newname = f"{year}_{abbrev(journal)}_{safe(authors,28)}_{safe(title,72)}.pdf"
        newname = re.sub(r"\s+", " ", newname)
        shutil.copy2(p, os.path.join(OUT, newname))
        dest = os.path.join(FT, "pdf", f"ID_{int(rec['no']):04d}.pdf")
        if not os.path.exists(dest):
            shutil.copy2(p, dest)
        log.append({"file": base, "matched": "YES", "how": how, "no": rec["no"],
                    "doi": rec["doi"], "newname": newname, "note": ""})

    with open(os.path.join(FT, "include_match_log.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["file", "matched", "how", "no", "doi", "newname", "note"])
        w.writeheader(); w.writerows(log)

    print(f"\n매칭 성공 {matched}/{len(pdfs)} · 실패 {len(unmatched)}")
    for b, d in unmatched:
        print(f"  ✗ {b} (doi={d})")
    print(f"리네임 사본: {OUT}")
    print(f"파이프라인 편입 후 fulltext/pdf 총 {len(glob.glob(os.path.join(FT,'pdf','*.pdf')))}개")


if __name__ == "__main__":
    main()
