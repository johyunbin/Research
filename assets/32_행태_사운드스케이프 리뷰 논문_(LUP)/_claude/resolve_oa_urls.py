# -*- coding: utf-8 -*-
"""
Paper32 — OA 32편의 DOI를 출판사 실제 URL로 해석하고 PDF 후보 URL을 만든다.
doi.org 리다이렉트만 따라가므로 봇 챌린지와 무관. 실제 내려받기는 브라우저로 한다.
출력: fulltext/oa_targets_resolved.csv
"""
import sys, os, csv, re, time, urllib.request, urllib.error

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) "
      "Chrome/125.0 Safari/537.36")


class NoRedirect(urllib.request.HTTPRedirectHandler):
    def redirect_request(self, req, fp, code, msg, headers, newurl):
        raise urllib.error.HTTPError(req.full_url, code, newurl, headers, fp)


def resolve(doi, hops=6):
    """doi.org → 최종 출판사 URL (리다이렉트 체인만 추적, 본문은 받지 않음)"""
    url = "https://doi.org/" + doi
    op = urllib.request.build_opener(NoRedirect)
    for _ in range(hops):
        req = urllib.request.Request(url, headers={"User-Agent": UA}, method="HEAD")
        try:
            op.open(req, timeout=25)
            return url
        except urllib.error.HTTPError as e:
            if e.code in (301, 302, 303, 307, 308):
                nxt = e.msg if isinstance(e.msg, str) and e.msg.startswith("http") else \
                    e.headers.get("Location", "")
                if not nxt:
                    return url
                url = urllib.parse.urljoin(url, nxt) if not nxt.startswith("http") else nxt
                continue
            return url          # 403/404 등도 최종 URL은 확보됨
        except Exception:
            return url
    return url


def pdf_candidates(final, doi, landing):
    """출판사별 PDF 직행 URL 후보"""
    f = final or ""
    out = []
    if "mdpi.com" in f:
        base = f.split("?")[0].rstrip("/")
        out += [base + "/pdf?version=1", base + "/pdf"]
    elif "frontiersin.org" in f:
        out += [f.rstrip("/") + "/pdf"]
    elif "sciencedirect.com" in f:
        m = re.search(r"/pii/([A-Z0-9]+)", f)
        if m:
            out += [f"https://www.sciencedirect.com/science/article/pii/{m.group(1)}/pdfft?"
                    f"isDTMRedir=true&download=true"]
    elif "aacus" in doi or "act-acustica" in f or "acta-acustica" in f:
        out += [f]
    elif "pubs.aip.org" in f or "asa.scitation.org" in f:
        out += [f]
    elif "journals.sagepub.com" in f:
        m = re.search(r"/doi/(?:abs/|full/)?(10\.\d+/[^?#]+)", f)
        if m:
            out += [f"https://journals.sagepub.com/doi/pdf/{m.group(1)}"]
    if landing and landing not in out:
        if "ncbi.nlm.nih.gov/pmc" in landing:
            m = re.search(r"(PMC\d+|/(\d{7,}))", landing)
            if m:
                pid = m.group(1) if m.group(1).startswith("PMC") else "PMC" + m.group(2)
                out += [f"https://www.ncbi.nlm.nih.gov/pmc/articles/{pid}/pdf/"]
        out += [landing]
    if f and f not in out:
        out += [f]
    seen, dedup = set(), []
    for u in out:
        if u and u not in seen:
            seen.add(u); dedup.append(u)
    return dedup


def main():
    import urllib.parse
    globals()["urllib"].parse = urllib.parse
    rows = list(csv.DictReader(open(os.path.join(FT, "oa_targets.csv"), encoding="utf-8-sig")))
    out = []
    for i, r in enumerate(rows, 1):
        fin = resolve(r["doi"])
        cands = pdf_candidates(fin, r["doi"], r["landing"])
        out.append({**r, "final_url": fin, "pdf_try1": cands[0] if cands else "",
                    "pdf_try2": cands[1] if len(cands) > 1 else "",
                    "pdf_try3": cands[2] if len(cands) > 2 else ""})
        host = re.sub(r"^https?://(www\.)?", "", fin).split("/")[0]
        print(f"  {r['rec']:>4} {host:<28} → {cands[0][:70] if cands else '(후보 없음)'}")
        time.sleep(0.5)
    with open(os.path.join(FT, "oa_targets_resolved.csv"), "w", newline="",
              encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(out[0].keys())); w.writeheader(); w.writerows(out)
    print(f"\n[저장] fulltext/oa_targets_resolved.csv ({len(out)}건)")


if __name__ == "__main__":
    import urllib.parse
    main()
