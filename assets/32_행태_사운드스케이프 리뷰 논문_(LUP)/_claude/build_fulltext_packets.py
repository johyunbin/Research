# -*- coding: utf-8 -*-
"""
Paper32 — 전문심사 패킷 생성
확보 PDF에서 심사·추출에 필요한 구간(초록/방법/결과 앞부분 + 표 캡션)을 압축 추출.
전문 통째가 아니라 판정·추출에 필요한 텍스트만 담아 에이전트 컨텍스트 절약.
출력: fulltext/packets/PKT_<no>.txt + fulltext/packet_index.csv
"""
import sys, os, re, csv, glob, fitz

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
PKT = os.path.join(FT, "packets")
os.makedirs(PKT, exist_ok=True)
MAXCH = 14000


def clean(t):
    t = re.sub(r"[ \t]+", " ", t)
    t = re.sub(r"\n{3,}", "\n\n", t)
    return t.strip()


def main():
    meta = {}
    with open(os.path.join(FT, "fulltext_status_20260803.csv"), encoding="utf-8-sig") as f:
        for r in csv.DictReader(f):
            meta[int(r["no"])] = r

    idx = []
    for p in sorted(glob.glob(os.path.join(FT, "pdf", "ID_*.pdf"))):
        no = int(os.path.basename(p)[3:7])
        m = meta.get(no, {})
        try:
            doc = fitz.open(p)
            full = "\n".join(doc[i].get_text() for i in range(len(doc)))
            npages = len(doc)
            doc.close()
        except Exception as e:
            print(f"  ERR {no}: {e}"); continue

        full = clean(full)
        head = full[:6000]                    # 제목·초록·서론 앞부분
        # 방법/결과 섹션 앵커
        body = ""
        for kw in ["Method", "METHOD", "Materials and method", "2. Method", "Study design",
                   "Data collection", "Participants", "Measures", "Procedure"]:
            i = full.find(kw, 3000)
            if i > 0:
                body = full[i:i + 6000]
                break
        if not body:
            body = full[6000:12000]
        # 표/그림 캡션(효과크기 단서)
        caps = re.findall(r"(?:Table|Fig(?:ure)?\.?)\s*\d+[^\n]{0,160}", full)[:25]

        txt = (f"### RECORD ID={no} | screening_verdict={m.get('verdict','?')} | year={m.get('year','?')}\n"
               f"JOURNAL: {m.get('journal','')}\nTITLE: {m.get('title','')}\nDOI: {m.get('doi','')}\n"
               f"PAGES: {npages}\n\n--- HEAD (title/abstract/intro) ---\n{head}\n\n"
               f"--- METHODS/RESULTS EXCERPT ---\n{body}\n\n--- TABLE/FIGURE CAPTIONS ---\n"
               + "\n".join(caps))
        txt = txt[:MAXCH]
        with open(os.path.join(PKT, f"PKT_{no:04d}.txt"), "w", encoding="utf-8") as f:
            f.write(txt)
        idx.append({"no": no, "verdict": m.get("verdict", ""), "year": m.get("year", ""),
                    "journal": m.get("journal", "")[:60], "title": m.get("title", "")[:90],
                    "chars": len(txt), "pages": npages})

    with open(os.path.join(FT, "packet_index.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=["no", "verdict", "year", "journal", "title", "chars", "pages"])
        w.writeheader(); w.writerows(idx)
    inc = sum(1 for r in idx if r["verdict"] == "INCLUDE")
    print(f"패킷 {len(idx)}개 생성 (INCLUDE {inc} · BORDERLINE {len(idx)-inc}) · 평균 {sum(r['chars'] for r in idx)//max(len(idx),1)}자")


if __name__ == "__main__":
    main()
