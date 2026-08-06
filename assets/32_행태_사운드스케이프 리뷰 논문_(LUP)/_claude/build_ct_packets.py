# -*- coding: utf-8 -*-
"""
Paper32 — 인용추적 확보 54편의 전문심사 패킷 생성
본검색 갈래(build_fulltext_packets.py)와 **동일한 기준·동일한 판정 양식**으로 심사하기 위한 입력.
전문 텍스트를 추출해 배치 파일로 나눈다(발췌가 아니라 전문 — 앞서 발췌 기반 한계를 겪었으므로).
출력: fulltext/ct_txt/CT####.txt · fulltext/ct_batches/ctft_XX.txt · ct_packet_index.csv
"""
import sys, os, csv, re

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
SRC = os.path.join(BASE, "ct_pdf")
TXT = os.path.join(FT, "ct_txt")
BAT = os.path.join(FT, "ct_batches")
os.makedirs(TXT, exist_ok=True)
os.makedirs(BAT, exist_ok=True)

try:
    import fitz
except ImportError:
    print("PyMuPDF 필요"); sys.exit(1)

PER_DOC = 46000          # 논문당 상한(장문 논문 꼬리 절단)
PER_BATCH = 6            # 배치당 논문 수


def clean(t):
    t = re.sub(r"[ \t]+", " ", t)
    t = re.sub(r"\n{3,}", "\n\n", t)
    return t.strip()


def main():
    fin = {r["rec"]: r for r in csv.DictReader(open(os.path.join(FT, "ct_screen_final.csv"),
                                                    encoding="utf-8-sig"))}
    files = sorted(f for f in os.listdir(SRC) if f.lower().endswith(".pdf"))
    idx, short = [], []

    for f in files:
        m = re.match(r"CT(\d{4})_", f)
        if not m:
            continue
        rec = str(int(m.group(1)))
        r = fin.get(rec, {})
        try:
            with fitz.open(os.path.join(SRC, f)) as d:
                pages = d.page_count
                txt = clean("\n".join(d[i].get_text() for i in range(pages)))
        except Exception as e:
            short.append((rec, f"읽기 실패 {type(e).__name__}")); continue
        if len(txt) < 3000:
            short.append((rec, f"텍스트 {len(txt)}자 — 스캔본 의심"))
        if len(txt) > PER_DOC:
            txt = txt[:PER_DOC] + "\n\n…[본문 이후 절단 — 필요 시 원문 PDF 참조]"
        open(os.path.join(TXT, f"CT{int(rec):04d}.txt"), "w", encoding="utf-8").write(txt)
        idx.append({"rec": rec, "screen": r.get("final", ""), "year": r.get("year", ""),
                    "journal": r.get("journal", ""), "doi": r.get("doi", ""),
                    "behavior_hint": r.get("behavior_hint", ""), "pages": pages,
                    "chars": len(txt), "title": r.get("title", ""), "pdf": f})

    # A군 먼저 배치(수확 높은 쪽을 앞에)
    idx.sort(key=lambda r: (r["screen"] != "RETRIEVE", int(r["rec"])))
    with open(os.path.join(FT, "ct_packet_index.csv"), "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(idx[0].keys())); w.writeheader(); w.writerows(idx)

    for p in os.listdir(BAT):
        os.remove(os.path.join(BAT, p))
    n = 0
    for i in range(0, len(idx), PER_BATCH):
        n += 1
        parts = []
        for r in idx[i:i + PER_BATCH]:
            body = open(os.path.join(TXT, f"CT{int(r['rec']):04d}.txt"), encoding="utf-8").read()
            parts.append(
                f"{'='*78}\n### REC {r['rec']} | 1차판정 {r['screen']} | {r['year']} | "
                f"{r['journal']}\nTITLE: {r['title']}\nDOI: {r['doi']}\n"
                f"초록단계 행태단서: {r['behavior_hint'] or '(없음 — 제목만으로 통과)'}\n"
                f"{'='*78}\n\n{body}\n")
        open(os.path.join(BAT, f"ctft_{n:02d}.txt"), "w", encoding="utf-8").write("\n\n".join(parts))

    from collections import Counter
    print(f"[완료] 텍스트 추출 {len(idx)}편 · 배치 {n}개(건당 {PER_BATCH})")
    print(f"  1차판정 구성: {dict(Counter(r['screen'] for r in idx))}")
    print(f"  평균 {sum(r['chars'] for r in idx)//len(idx):,}자 · 최소 {min(r['chars'] for r in idx):,}자")
    if short:
        print(f"  ⚠️ 확인 필요 {len(short)}건: {short}")


if __name__ == "__main__":
    main()
