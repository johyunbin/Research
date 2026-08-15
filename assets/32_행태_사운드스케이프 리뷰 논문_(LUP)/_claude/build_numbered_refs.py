# -*- coding: utf-8 -*-
"""
Paper32 — 저자·연도 인용 → **번호 인용 [N]** 전환 + Vancouver 서지 생성

서식 기준은 paper31 `Manuscript_EN_20260727_005544.docx` 실측:
  본문 : `... 라고 보고했다 [17]` · 연속은 `[17,18,29]`
  서지 : `[N] Family I, Family I, ..., et al. Title. Journal. Year;Vol:pages. doi:10.x/y`
         저자 6인 초과 시 6인 + `et al.` · 이니셜에 마침표 없음 · doi 는 소문자 prefix

번호는 **본문 첫 등장 순서**로 매긴다(Vancouver 규칙).
출력: 원고 in-place 수정 + fulltext/references_numbered.md
"""
import sys, os, re, json, time
import urllib.parse, urllib.request

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
MS = os.path.join(os.path.dirname(BASE), "01_논문작업", "Manuscript_KO_20260806.md")
UA = {"User-Agent": "paper32-SR/0.1 (mailto:wh8502@naver.com)"}

# ── 본문 인용 표기 → 문헌 키 ────────────────────────────────────
# 정규식이 아니라 **문자열 그대로** 매칭한다. 긴 것부터 치환해야 부분 치환 사고가 없다.
CITE = {
    "Whyte(1980)":                                      ["whyte1980"],
    "Gehl(2011/1971)":                                  ["gehl2011"],
    "(Miedema & Oudshoorn, 2001)":                      ["miedema2001"],
    "(Basner et al., 2014)":                            ["basner2014"],
    "Schafer(1994/1977)":                               ["schafer1994"],
    "(International Organization for Standardization, 2014)": ["iso2014"],
    "Axelsson 등(2010)":                                 ["axelsson2010"],
    "Aletta 등(2016)":                                   ["aletta2016"],
    "Kang 등(2016)":                                     ["kang2016"],
    "(Zhang et al., 2025)":                             ["zhang2025"],
    "Wang 등(2026)":                                     ["wang2026"],
    "Zhang 등(2025)":                                    ["zhang2025"],
    "Buxton 등(2021)":                                   ["buxton2021"],
    "(Page et al., 2021)":                              ["page2021"],
    "(Hong et al., 2018)":                              ["hong2018"],
    "(Hedges, 1981)":                                   ["hedges1981"],
    "(Chinn, 2000)":                                    ["chinn2000"],
    "(Hartung & Knapp, 2001)":                          ["hartung2001"],
    "(Higgins et al., 2009)":                           ["higgins2009"],
    "Moser(1988)":                                      ["c14"],
    "Mathews와 Canon(1975)":                             ["cCT0025"],
    "Chen과 Kang(2023)":                                 ["c931"],
    "Chen 등(2024)":                                     ["c1069"],
    "(Franěk 등, 2014)":                                 ["cCT0090"],
    "(Franěk & Režný, 2021)":                            ["cCT0166"],
    "(Schrapel 등, 2022)":                               ["cCT0335"],
    "(Kang et al., 2016; Aletta et al., 2016)":         ["kang2016", "aletta2016"],
    "(Franěk 등, 2018, 2019)":                           ["c461", "c532"],
    "(Berkouk 등, 2020)":                                ["c617"],
    "Ba 등(2024)":                                       ["c941"],
    "(Aletta 등, 2016)":                                 ["c323"],   # 코퍼스 연구로서의 Aletta 2016
    "(Montes González 등, 2023)":                        ["cCT0126"],
    "(Buxton et al., 2021)":                            ["buxton2021"],
    "(Dzhambov 등, 2026; Huang 등, 2023)":                ["c1226", "c951"],
}

# 배경 문헌 DOI (전건 Crossref 확인분)
BG_DOI = {
    "aletta2016": "10.1016/j.landurbplan.2016.02.001",
    "axelsson2010": "10.1121/1.3493436",
    "kang2016": "10.1016/j.buildenv.2016.08.011",
    "wang2026": "10.1080/23748834.2026.2683259",
    "zhang2025": "10.1016/j.landurbplan.2025.105463",
    "page2021": "10.1136/bmj.n71",
    "higgins2009": "10.1111/j.1467-985X.2008.00552.x",
    "basner2014": "10.1016/S0140-6736(13)61613-X",
    "chinn2000": "10.1002/1097-0258(20001130)19:22<3127::aid-sim784>3.0.co;2-m",
    "hartung2001": "10.1002/sim.791",
    "hedges1981": "10.3102/10769986006002107",
    "hong2018": "10.3233/EFI-180221",
    "buxton2021": "10.1073/pnas.2013097118",
    "miedema2001": "10.1289/ehp.01109409",
}
# DOI 가 없는 1차 출처(표준·단행본)
AUTHOR_FIX = {   # Crossref 에 저자가 비어 있는 레코드 — 원문 표지에서 보완
    "c941": "Ba M, Li Z, Kang J",
}
BG_MANUAL = {
    "iso2014": "International Organization for Standardization. ISO 12913-1:2014 Acoustics — "
               "Soundscape — Part 1: Definition and conceptual framework. Geneva: ISO; 2014.",
    "schafer1994": "Schafer RM. The Soundscape: Our Sonic Environment and the Tuning of the World. "
                   "Rochester (VT): Destiny Books; 1994.",
    "gehl2011": "Gehl J. Life Between Buildings: Using Public Space. Washington (DC): "
                "Island Press; 2011.",
    "whyte1980": "Whyte WH. The Social Life of Small Urban Spaces. Washington (DC): "
                 "Conservation Foundation; 1980.",
}

MOJI = re.compile(r"[ÃÄÅÂÐ×Þ][^\x00-\x7f]")


def demoji(s):
    if not s or not MOJI.search(s):
        return s
    for e in ("cp1252", "latin-1"):
        try:
            f = s.encode(e).decode("utf-8")
        except Exception:
            continue
        if not MOJI.search(f) and "�" not in f:
            return f
    return s


def cr(doi):
    u = "https://api.crossref.org/works/" + urllib.parse.quote(doi)
    for i in range(3):
        try:
            with urllib.request.urlopen(urllib.request.Request(u, headers=UA), timeout=45) as r:
                return json.loads(r.read().decode("utf-8"))["message"]
        except Exception:
            if i < 2:
                time.sleep(2 * (i + 1))
    return None


def vancouver(m):
    """paper31 실측 서식: Family I, Family I, et al. Title. Journal. Year;Vol:pages. doi:..."""
    au = m.get("author") or []
    names = []
    for a in au:
        fam = demoji(a.get("family") or a.get("name") or "")
        giv = demoji(a.get("given") or "")
        ini = "".join(p[0] for p in re.split(r"[\s\-]+", giv) if p)
        names.append(f"{fam} {ini}".strip())
    who = ", ".join(names[:6]) + (", et al." if len(names) > 6 else "")
    yr = ""
    for k in ("published-print", "published-online", "issued"):
        p = (m.get(k) or {}).get("date-parts") or []
        if p and p[0] and p[0][0]:
            yr = str(p[0][0]); break
    ti = demoji((m.get("title") or [""])[0]).rstrip(".")
    ct = demoji((m.get("container-title") or [""])[0]).replace("&amp;", "&")
    s = f"{who.rstrip('.')}. {ti}. {ct}. {yr}"
    if m.get("volume"):
        s += f";{m['volume']}"
        if m.get("page"):
            s += f":{m['page']}"
    elif m.get("page"):
        s += f":{m['page']}"
    s += f". doi:{(m.get('DOI') or '').lower()}"
    return s


def main():
    t = open(MS, encoding="utf-8").read()

    # ── 코퍼스 문헌 DOI(uid → doi) ────────────────────────────────
    import csv
    cdoi = {}
    for r in csv.DictReader(open(os.path.join(FT, "references.csv"), encoding="utf-8-sig")):
        if r.get("doi"):
            cdoi["c" + r["uid"]] = r["doi"]

    need = sorted({k for ks in CITE.values() for k in ks})
    missing = [k for k in need if k not in BG_DOI and k not in BG_MANUAL and k not in cdoi]
    if missing:
        raise SystemExit(f"⚠️ 서지 출처 없는 인용 키: {missing}")

    # ── 등장 순서로 번호 부여 ────────────────────────────────────
    body_end = t.index("## References")
    order, num = [], {}
    hits = []
    for lit, keys in CITE.items():
        start = 0
        while True:
            i = t.find(lit, start)
            if i < 0 or i >= body_end:
                break
            hits.append((i, lit, keys)); start = i + 1
    hits.sort()
    for _, _, keys in hits:
        for k in keys:
            if k not in num:
                num[k] = len(order) + 1
                order.append(k)

    # ── 본문 치환(긴 표기부터) ───────────────────────────────────
    body, tail = t[:body_end], t[body_end:]
    for lit in sorted(CITE, key=len, reverse=True):
        tag = "[" + ",".join(str(num[k]) for k in CITE[lit]) + "]"
        # "Whyte(1980)는" → "Whyte [1]는" / "(Basner et al., 2014)" → " [4]"
        if lit.startswith("("):
            body = body.replace(" " + lit, " " + tag).replace(lit, tag)
        else:
            name = re.sub(r"\(.*", "", lit).strip()
            body = body.replace(lit, f"{name} {tag}")
    t = body + tail

    # ── 서지 생성 ────────────────────────────────────────────────
    lines, fails = [], []
    for k in order:
        if k in BG_MANUAL:
            lines.append(f"[{num[k]}] {BG_MANUAL[k]}"); continue
        doi = BG_DOI.get(k) or cdoi.get(k)
        m = cr(doi)
        if not m:
            fails.append((k, doi)); lines.append(f"[{num[k]}] ⚠️ 서지 미확보 (doi:{doi})")
            continue
        v = vancouver(m)
        if k in AUTHOR_FIX:
            v = AUTHOR_FIX[k] + ". " + v.lstrip(". ")
        lines.append(f"[{num[k]}] {v}")
        time.sleep(0.35)

    ref = ("## References\n\n"
           "> 번호는 본문 첫 등장 순서다. 포함 98편의 전건 목록은 Supplementary S5, "
           "전문 단계 배제 문헌과 사유는 S6에 있다. 배경 문헌은 모두 Crossref에서 서지를 "
           "확인했다.\n\n" + "\n\n".join(lines) + "\n\n")
    j = t.index("---\n\n## Supplementary material")
    t = t[:t.index("## References")] + ref + t[j:]
    open(MS, "w", encoding="utf-8").write(t)
    open(os.path.join(FT, "references_numbered.md"), "w", encoding="utf-8").write(
        "\n\n".join(lines) + "\n")

    print(f"[완료] 인용 {len(hits)}개소 → 번호 {len(order)}건")
    print(f"       본문 [N] 표기 {len(re.findall(r'\\[\\d+(?:,\\d+)*\\]', t[:t.index('## References')]))}개")
    if fails:
        print(f"⚠️ 서지 미확보 {len(fails)}: {fails}")
    print("[저장] 원고 in-place · fulltext/references_numbered.md")


if __name__ == "__main__":
    main()
