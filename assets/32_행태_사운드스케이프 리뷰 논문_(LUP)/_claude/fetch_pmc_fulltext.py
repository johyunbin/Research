# -*- coding: utf-8 -*-
"""
Paper32 — 전문 수집 2단계: PubMed Central OA 서비스 경유 (공식 대량수집 경로)
DOI→PMCID(idconv) → oa.fcgi(OA 서브셋 링크) → PDF 다운로드.
이미 확보(ID_*.pdf 존재)는 건너뜀. 출력: fulltext/pmc_log_*.csv
"""
import sys, csv, os, re, json, time, urllib.parse, urllib.request, urllib.error

sys.stdout.reconfigure(encoding="utf-8")
TIMECODE = "20260802_221808"
BASE = r"C:\Users\wh850\Research\assets\32_행태_사운드스케이프 리뷰 논문_(LUP)\_claude"
FT = os.path.join(BASE, "fulltext")
UA = "paper32-systematic-review/0.1 (academic use; contact via OSF osf.io/7ew8q)"


def get(url, timeout=120, tries=4):
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    for i in range(tries):
        try:
            return urllib.request.urlopen(req, timeout=timeout)
        except (urllib.error.HTTPError, urllib.error.URLError) as e:
            code = getattr(e, "code", None)
            if i < tries - 1 and code in (None, 429, 500, 502, 503):
                time.sleep(4 * (i + 1)); continue
            raise


def main():
    # 미확보 대상 로드 (oa_log에서 OK/ALREADY 아닌 것)
    with open(os.path.join(FT, f"oa_log_{TIMECODE}.csv"), encoding="utf-8-sig") as f:
        rows = list(csv.DictReader(f))
    have = {r["no"] for r in rows if r["status"] in ("OK", "ALREADY")}
    todo = [r for r in rows if r["no"] not in have and r["doi"]]
    print(f"[1] 잔여 {len(todo)}건 (DOI 보유)")

    # DOI → PMCID (idconv, 100건 배치)
    pmcid = {}
    for i in range(0, len(todo), 100):
        batch = todo[i:i+100]
        ids = ",".join(urllib.parse.quote(r["doi"], safe="") for r in batch)
        url = f"https://www.ncbi.nlm.nih.gov/pmc/utils/idconv/v1.0/?tool=paper32sr&format=json&ids={ids}"
        try:
            data = json.loads(get(url).read().decode("utf-8"))
            for rec in data.get("records", []):
                d = (rec.get("doi") or "").lower()
                p = rec.get("pmcid")
                if d and p:
                    pmcid[d] = p
        except Exception as e:
            print(f"    idconv 배치 {i//100+1} 실패: {e}")
        time.sleep(1.0)
    print(f"[2] PMCID 매핑 {len(pmcid)}건")

    # oa.fcgi → PDF 링크 → 다운로드
    ok, log = 0, []
    for idx, r in enumerate(todo, 1):
        p = pmcid.get(r["doi"].lower())
        status = ""
        if not p:
            status = "NO_PMCID"
        else:
            try:
                xml = get(f"https://www.ncbi.nlm.nih.gov/pmc/utils/oa/oa.fcgi?id={p}").read().decode("utf-8", "ignore")
                m = re.search(r'href="([^"]+\.pdf)"', xml)
                if not m:
                    status = "NO_OA_PDF(tgz-only or not in OA subset)"
                else:
                    href = m.group(1).replace("ftp://ftp.ncbi.nlm.nih.gov", "https://ftp.ncbi.nlm.nih.gov")
                    data = get(href, timeout=180).read()
                    if data[:5] == b"%PDF-":
                        dest = os.path.join(FT, "pdf", f"ID_{int(r['no']):04d}.pdf")
                        with open(dest, "wb") as f:
                            f.write(data)
                        status, ok = "OK_PMC", ok + 1
                    else:
                        status = "NOT_PDF"
            except Exception as e:
                status = f"ERR:{type(e).__name__}"
            time.sleep(1.2)  # NCBI 3req/s 준수 여유
        log.append({"no": r["no"], "doi": r["doi"], "pmcid": p or "", "status": status})
        if idx % 25 == 0:
            print(f"    진행 {idx}/{len(todo)} (PMC 성공 {ok})")

    with open(os.path.join(FT, f"pmc_log_{TIMECODE}.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["no", "doi", "pmcid", "status"])
        w.writeheader(); w.writerows(log)
    from collections import Counter
    print(f"\n[완료] PMC 추가 확보 {ok}건 → 누적 {len(have)+ok}/189")
    print("상태:", dict(Counter(l['status'] for l in log)))


if __name__ == "__main__":
    main()
