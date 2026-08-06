# -*- coding: utf-8 -*-
"""
Paper32 — 사용자가 받아준 PDF를 인용추적 REC 번호에 대조·편입
파일명은 무엇이든 상관없다. 본문 앞부분 텍스트와 DOI를 뽑아 ct_screen_final.csv의 제목·DOI와 매칭한다.
사용: python match_ct_pdfs.py <검색할 폴더> [...]
출력: ct_pdf/CT####_연도_제목.pdf · fulltext/ct_match_log.csv
"""
import sys, os, csv, re, shutil, unicodedata, difflib

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
DEST = os.path.join(BASE, "ct_pdf")
os.makedirs(DEST, exist_ok=True)

try:
    import fitz                                   # PyMuPDF
except ImportError:
    print("PyMuPDF 필요: pip install pymupdf"); sys.exit(1)


def norm(s):
    s = unicodedata.normalize("NFKD", (s or "").lower())
    return re.sub(r"[^a-z0-9]", "", s)


def safe(s, n=88):
    s = re.sub(r"[^\w\s\-.,()]", "", s, flags=re.UNICODE)
    return re.sub(r"\s+", " ", s).strip()[:n]


def head_text(path, pages=2, cap=6000):
    try:
        with fitz.open(path) as d:
            return " ".join(d[i].get_text() for i in range(min(pages, d.page_count)))[:cap]
    except Exception as e:
        return f"__ERR__{type(e).__name__}"


def main():
    dirs = sys.argv[1:] or [os.path.dirname(BASE)]
    tg = [r for r in csv.DictReader(open(os.path.join(FT, "ct_screen_final.csv"),
                                         encoding="utf-8-sig"))
          if r["final"] in ("RETRIEVE", "UNCERTAIN")]
    by_doi = {(r["doi"] or "").lower(): r for r in tg if r["doi"]}
    titles = [(norm(r["title"]), r) for r in tg]
    print(f"[0] 대조 대상 {len(tg)}건 · 검색 폴더 {dirs}")

    # 이미 편입된 REC
    done = set()
    for f in os.listdir(DEST):
        m = re.match(r"CT(\d{4})_", f)
        if m:
            done.add(str(int(m.group(1))))
    print(f"[1] 이미 편입 {len(done)}건")

    cands = []
    for d in dirs:
        if not os.path.isdir(d):
            print(f"    ⚠️ 폴더 없음: {d}"); continue
        for f in sorted(os.listdir(d)):
            if f.lower().endswith(".pdf"):
                cands.append(os.path.join(d, f))
    print(f"[2] PDF 후보 {len(cands)}개")

    log, moved, skipped = [], 0, 0
    for p in cands:
        txt = head_text(p)
        if txt.startswith("__ERR__"):
            log.append({"file": os.path.basename(p), "rec": "", "how": txt, "score": ""})
            print(f"    ⚠️ 읽기 실패 {os.path.basename(p)} — {txt}"); continue
        low = txt.lower()

        hit, how, score = None, "", 0.0
        # ① DOI 직매치
        for m in re.finditer(r"10\.\d{4,9}/[^\s\"'<>,;)\]]+", low):
            d = m.group(0).rstrip(".,;)")
            for cut in range(len(d), 6, -1):
                if d[:cut] in by_doi:
                    hit, how, score = by_doi[d[:cut]], "doi", 1.0
                    break
            if hit:
                break
        # ② 제목 유사도
        if not hit:
            nt = norm(txt[:1800])
            best, bs = None, 0.0
            for tn, r in titles:
                if len(tn) < 20:
                    continue
                if tn[:60] in nt:
                    best, bs = r, 0.99
                    break
                s = difflib.SequenceMatcher(None, tn[:120], nt[:900]).ratio()
                if s > bs:
                    best, bs = r, s
            if best and bs >= 0.42:
                hit, how, score = best, "title", bs

        base = os.path.basename(p)
        if not hit:
            skipped += 1
            log.append({"file": base, "rec": "", "how": "no-match", "score": ""})
            print(f"    ? 매칭 실패 {base}")
            continue
        rec = hit["rec"]
        if rec in done:
            log.append({"file": base, "rec": rec, "how": how + "/dup", "score": f"{score:.2f}"})
            print(f"    = REC {rec:>4} 이미 있음 ({base})")
            continue
        fn = f"CT{int(rec):04d}_{hit.get('year','')}_{safe(hit['title'])}.pdf"
        shutil.copy2(p, os.path.join(DEST, fn))
        done.add(rec); moved += 1
        log.append({"file": base, "rec": rec, "how": how, "score": f"{score:.2f}",
                    "title": hit["title"], "dest": fn})
        print(f"    ✓ REC {rec:>4} [{how} {score:.2f}] {hit['title'][:58]}")

    with open(os.path.join(FT, "ct_match_log.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["file", "rec", "how", "score", "title", "dest"],
                           extrasaction="ignore")
        w.writeheader(); w.writerows(log)

    # 남은 대상
    remain = [r for r in tg if r["rec"] not in done]
    print(f"\n[완료] 신규 편입 {moved} · 매칭 실패 {skipped} · ct_pdf 총 {len(done)}건")
    print(f"       남은 확보 대상 {len(remain)}건 "
          f"(A {sum(1 for r in remain if r['final']=='RETRIEVE')} · "
          f"B {sum(1 for r in remain if r['final']=='UNCERTAIN')})")


if __name__ == "__main__":
    main()
