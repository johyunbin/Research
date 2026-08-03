# -*- coding: utf-8 -*-
"""
Paper32 — 벤치마크 리콜 체크 (DOI 정밀판, 429 백오프 포함)
파일럿 시드 논문 5편이 본검색식(A AND B AND C + 필터)에 걸리는지 DOI 단위로 검증.
MISS면 어느 조건(쿼리 vs 필터)에서 탈락했는지 진단.
"""
import sys, json, time, urllib.parse, urllib.request, urllib.error

sys.stdout.reconfigure(encoding="utf-8")
BASE = "https://api.openalex.org/works"

from openalex_pilot_search import QUERY, FILTERS  # 동일 검색식 재사용

BENCHMARKS = [
    ("Franek 2018 보행속도",   "10.3390/ijerph15040752"),
    ("Aletta 2016 공공공간 실험", "10.3390/app6100276"),
    ("Steele 2019 Musikiosk",  "10.3390/ijerph16101865"),
    ("Meng 2018 군중음악",      "10.3389/fpsyg.2018.00596"),
    ("Song 2020 LUP 공원행태",  "10.1016/j.landurbplan.2019.103890"),
]


def fetch(params: dict, tries: int = 5) -> dict:
    url = BASE + "?" + urllib.parse.urlencode(params)
    req = urllib.request.Request(url, headers={"User-Agent": "paper32-pilot/0.2"})
    for i in range(tries):
        try:
            with urllib.request.urlopen(req, timeout=60) as r:
                return json.loads(r.read().decode("utf-8"))
        except urllib.error.HTTPError as e:
            if e.code == 429 and i < tries - 1:
                wait = 5 * (i + 1)
                print(f"      ... 429 → {wait}s 대기 후 재시도")
                time.sleep(wait)
                continue
            raise


def count(filt: str) -> int:
    return fetch({"filter": filt, "per-page": 1})["meta"]["count"]


def main():
    hits = 0
    for label, doi in BENCHMARKS:
        exists = count(f"doi:{doi}")
        if exists == 0:
            print(f"[{label}] ⚠️ OpenAlex에 DOI 레코드 없음 ({doi})")
            time.sleep(1.5)
            continue
        q_hit = count(f"doi:{doi},title_and_abstract.search:{QUERY}")
        full_hit = count(f"doi:{doi},title_and_abstract.search:{QUERY},{FILTERS}")
        if full_hit:
            hits += 1
            print(f"[{label}] HIT — 검색식+필터 모두 통과")
        elif q_hit:
            # 필터에서 탈락 — 어느 필터인지 분해
            lang = count(f"doi:{doi},title_and_abstract.search:{QUERY},language:en")
            typ = count(f"doi:{doi},title_and_abstract.search:{QUERY},type:article")
            src = count(f"doi:{doi},title_and_abstract.search:{QUERY},primary_location.source.type:journal")
            print(f"[{label}] MISS(필터) — language:en={lang} type:article={typ} source:journal={src}")
        else:
            # 쿼리 블록에서 탈락 — 블록별 분해
            from openalex_pilot_search import BLOCK_A, BLOCK_B, BLOCK_C
            a = count(f"doi:{doi},title_and_abstract.search:{BLOCK_A}")
            b = count(f"doi:{doi},title_and_abstract.search:{BLOCK_B}")
            c = count(f"doi:{doi},title_and_abstract.search:{BLOCK_C}")
            print(f"[{label}] MISS(쿼리) — BlockA={a} BlockB={b} BlockC={c}")
        time.sleep(1.5)
    print(f"\n리콜: {hits}/{len(BENCHMARKS)}")


if __name__ == "__main__":
    main()
