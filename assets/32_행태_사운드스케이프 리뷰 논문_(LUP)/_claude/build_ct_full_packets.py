# -*- coding: utf-8 -*-
"""
Paper32 — 인용추적 포함분 전문 재추출 (절단 없이)
1차 패킷은 논문당 46,000자 상한을 뒀는데, 포함 16편 중 13편이 그 한도를 넘겨 잘렸다.
MMAT는 '보고가 없으면 CT'로 판정하는 도구라 **추출 한도가 품질 점수를 깎는 인공물**이 된다.
여기서는 상한 없이 다시 뽑고, 재평가 대상(절단됐던 논문)만 배치로 만든다.
출력: fulltext/ct_txt_full/CT####.txt · fulltext/ct_batches_full/ctfull_XX.txt
"""
import sys, os, csv, re

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
SRC = os.path.join(BASE, "ct_pdf")
TXT = os.path.join(FT, "ct_txt_full")
BAT = os.path.join(FT, "ct_batches_full")
os.makedirs(TXT, exist_ok=True)
os.makedirs(BAT, exist_ok=True)

import fitz

OLD_CAP = 46000
PER_BATCH = 3          # 전문이 길어졌으므로 배치당 논문 수를 줄인다


def clean(t):
    t = re.sub(r"[ \t]+", " ", t)
    return re.sub(r"\n{3,}", "\n\n", t).strip()


def main():
    inc = {r["no"]: r for r in csv.DictReader(open(os.path.join(FT, "ct_extraction_final.csv"),
                                                   encoding="utf-8-sig"))}
    idx = {r["rec"]: r for r in csv.DictReader(open(os.path.join(FT, "ct_packet_index.csv"),
                                                    encoding="utf-8-sig"))}
    pdfs = {}
    for f in os.listdir(SRC):
        m = re.match(r"CT(\d{4})_", f)
        if m:
            pdfs[str(int(m.group(1)))] = f

    rows = []
    for n in sorted(inc, key=int):
        f = pdfs.get(n)
        if not f:
            print(f"  ⚠️ REC {n} PDF 없음"); continue
        with fitz.open(os.path.join(SRC, f)) as d:
            pages = d.page_count
            txt = clean("\n".join(d[i].get_text() for i in range(pages)))
        open(os.path.join(TXT, f"CT{int(n):04d}.txt"), "w", encoding="utf-8").write(txt)
        rows.append({"rec": n, "chars": len(txt), "pages": pages,
                     "was_truncated": "yes" if len(txt) > OLD_CAP else "no",
                     "title": inc[n]["title"], "journal": inc[n]["journal"],
                     "year": inc[n]["year"], "verdict": inc[n]["final_verdict"]})

    with open(os.path.join(FT, "ct_full_text_index.csv"), "w", newline="",
              encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=list(rows[0].keys())); w.writeheader(); w.writerows(rows)

    # 재평가 대상 = 1차에서 절단됐던 논문
    redo = [r for r in rows if r["was_truncated"] == "yes"]
    redo.sort(key=lambda r: -r["chars"])       # 긴 것부터 고르게 분배
    buckets = [[] for _ in range((len(redo) + PER_BATCH - 1) // PER_BATCH)]
    for i, r in enumerate(redo):               # 라운드로빈으로 길이 균형
        buckets[i % len(buckets)].append(r)

    for p in os.listdir(BAT):
        os.remove(os.path.join(BAT, p))
    for bi, b in enumerate(buckets, start=1):
        parts = []
        for r in b:
            body = open(os.path.join(TXT, f"CT{int(r['rec']):04d}.txt"), encoding="utf-8").read()
            parts.append(f"{'='*78}\n### REC {r['rec']} | {r['verdict']} | {r['year']} | "
                         f"{r['journal']}\nTITLE: {r['title']}\n"
                         f"({r['pages']}쪽 · {len(body):,}자 · 절단 없음)\n{'='*78}\n\n{body}\n")
        open(os.path.join(BAT, f"ctfull_{bi:02d}.txt"), "w",
             encoding="utf-8").write("\n\n".join(parts))

    print(f"[완료] 전문 재추출 {len(rows)}편 (절단 없음)")
    print(f"  1차에서 절단됐던 논문 {len(redo)}편 → 재평가 배치 {len(buckets)}개")
    for bi, b in enumerate(buckets, start=1):
        print(f"    ctfull_{bi:02d}: {[r['rec'] for r in b]} "
              f"({sum(r['chars'] for r in b):,}자)")
    print(f"  최장 {max(r['chars'] for r in rows):,}자 · 평균 {sum(r['chars'] for r in rows)//len(rows):,}자")


if __name__ == "__main__":
    main()
