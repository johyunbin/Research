# -*- coding: utf-8 -*-
"""
Paper32 — 추가 전문평가 판정을 코퍼스 정본에 반영하는 재실행 가능한 통합 스크립트 (2026-09-18)

입력(여러 세트 가능): fulltext/newft_final[_<날짜>]/{verdict,extract,mmat,detail,es}.csv
  - verdict.csv : id, branch(DB/CT/OAS), verdict, reason_code(X1–X7), boundary_rule, confidence, rationale …
  - extract·mmat·detail : 포함·SENS 만 (detail = 연구당 MMAT 5문항, item_no 필수)
서지: fulltext/newft_manifest_*.csv (id 별 year·journal·title·doi·screening; 뒤 파일이 앞 파일을 덮음)
8월 코퍼스 판정 변경: 아래 CORPUS_CHANGES (+ 선택 --changes CSV, 같은 열)

반영 위치 (review/pipeline_map_add_studies_20260917.md §1–2)
  DB : ft_verdicts_v2.csv(모든 판정) · ft_extraction_v2.csv · quality_all.csv · quality_detail_all.csv(포함·SENS)
  CT : ct_verdicts_final.csv · ct_extraction_final.csv · qa_results/qa_ctfull_<세트>_{mmat,detail}.csv
       (merge_quality_v2.py 가 qa_ctfull_* 을 glob 으로 읽는다) · ct_retrieval_final.csv state → retrieved-<날짜>
  OAS: oas_results_ft/oasft_<세트>_{verdict,extract}.csv (finalize_oa_supp_branch.py 가 glob 으로 읽도록 수정)
       · qa_results/qa_oas_{mmat,detail}.csv(행 추가) · oa_supp_retrieval.csv result → ok-<날짜>
  공통: ruling_audit.csv(8월 판정 변경 기록) · newft_integration_ledger.csv(멱등성 원장)
        ft_exclusion_reasons.csv(3경로 전문 배제 전건의 사유 범주 — fig_data.py PRISMA 가 센다, 매번 재생성)
        retrieval_status_all.csv(3경로 전문 확보 대상 전건의 확보 상태·미확보 사유 — fig_data.py 가 센다, 매번 재생성)
        appendix_b_author_lookup.json · appendix_b_crossref_all.json (--fetch-authors: 새 포함 연구 저자 Crossref 조회)

멱등성
  - 원장에 같은 내용 해시로 기록된 id 는 건너뛴다(재실행해도 중복 행이 생기지 않는다).
  - 원장에 없는데 대상 파일에 id 가 이미 있거나, 원장 해시와 내용이 다르면 **아무것도 쓰지 않고 중단**한다.
  - 파생 표 2종은 정본에서 매번 다시 만든다.

실행
  python integrate_newft_into_corpus.py fulltext/newft_final                      # 검사·계획만 (쓰기 없음)
  python integrate_newft_into_corpus.py fulltext/newft_final --apply              # 반영
  python integrate_newft_into_corpus.py fulltext/newft_final fulltext/newft_final_20260919 --apply --fetch-authors --downstream
  --downstream 만 따로: 반영 없이 다운스트림 재생성 (지도 §3 순서, 메타분석 스크립트 제외)
  --ma-candidates out.csv : es.csv 에서 기존 4 클러스터 후보만 뽑아 저장(메타분석 입력은 건드리지 않는다)
"""
import argparse, csv, glob, hashlib, json, os, re, subprocess, sys, time, urllib.parse, urllib.request
from collections import Counter, defaultdict
from datetime import datetime, timezone, timedelta

sys.stdout.reconfigure(encoding="utf-8")
BASE = os.path.dirname(os.path.abspath(__file__))
FT = os.path.join(BASE, "fulltext")
QA = os.path.join(FT, "qa_results")
OASFT = os.path.join(FT, "oas_results_ft")
KST = timezone(timedelta(hours=9))
NOW = datetime.now(KST).strftime("%Y-%m-%d %H:%M KST")

KEEP = ("FINAL_INCLUDE", "SENS_ONLY")
VERDICTS = ("FINAL_INCLUDE", "SENS_ONLY", "FINAL_EXCLUDE")
ID_RE = {"DB": re.compile(r"^\d+$"), "CT": re.compile(r"^CT\d{4}$"), "OAS": re.compile(r"^OAS\d{4}$")}
RULES = {"", "R1", "R2", "R3", "P1", "P2", "P3"}
EXT_COLS = ["no", "final_verdict", "year", "journal", "title", "country", "setting", "design",
            "sample_n", "exposure", "behaviour_domain", "behaviour_measure",
            "measurement_method", "direction", "key_finding", "effect_stats"]
Q_ITEMS = ("Q1", "Q2", "Q3", "Q4", "Q5")
CAT_PREFIX = {"qualitative": "1", "quantitative rct": "2", "quantitative non-randomised": "3",
              "quantitative descriptive": "4", "mixed methods": "5"}

# ── X 코드 → PRISMA 전문 배제 범주 (fig_data.py 가 이 category 열을 센다) ─────────────────────
#   X2→세팅, X3→노출, X4→관찰 가능한 행태 결과 없음, X5→비실증, X6→학술지 논문 아님(새 범주), X7→언어
X_CATEGORY = {
    "X1": "Animal",
    "X2": "Setting not eligible",
    "X3": "No acoustic exposure",
    "X4": "No observed behaviour",
    "X5": "Not empirical",
    "X6": "Not a journal article",
    "X7": "Not in English",
}
# 8월 흐름도(fig_data.py·Fig. 1)의 경로별 표기를 그대로 잇는다 — 같은 사유, 표기만 다름
X_CATEGORY_BRANCH = {("CT", "X3"): "No acoustic variable", ("OAS", "X3"): "No acoustic variable",
                     ("OAS", "X4"): "No behavioural outcome"}
X_KO = {"X1": "동물·야생", "X2": "세팅 부적합", "X3": "음환경 노출 없음", "X4": "관찰 가능한 행태 결과 없음",
        "X5": "비실증", "X6": "학술지 논문 아님", "X7": "영어 본문 아님"}

# 8월 DB 전문 배제 16건 — 코드 기록이 없고 집계만 남아 있다(prisma_flow.md: 세팅 7·행태 6·지각 1·노출 1·비실증 1).
# ft_verdicts_v2.reason 서술로 재배정해 집계를 그대로 재현한다. ⚠️ 234(세팅↔행태)·909(행태↔비실증)는 해석 여지 —
# 909·1243 둘 다 사유가 비실증인데 원 집계는 비실증 1건이라, 원 집계 보존을 위해 909 를 X4 로 둔다(보고서에 기재).
AUG_DB_EXCLUSION = {
    "234": "X2", "613": "X2", "779": "X2", "880": "X2", "912": "X2", "1141": "X2", "1154": "X2",
    "648": "X4", "770": "X4", "811": "X4", "909": "X4", "981": "X4", "1253": "X4",
    "948": "PERCEPTUAL", "729": "X3", "1243": "X5",
}
SPECIAL_CATEGORY = {"PERCEPTUAL": "Perceptual outcome"}

# 8월 코퍼스 판정 변경 (사용자 위임 2026-09-18: 권장안 채택)
CORPUS_CHANGES = [
    {"uid": "1226", "branch": "DB", "old": "FINAL_INCLUDE", "new": "FINAL_EXCLUDE", "rule": "R1", "reason_code": "X2",
     "reason": "경계판정① 주거노출·공간비특정 PA — 세팅 기준 위반 (2026-09-18 재판정 R1: 779·880·1141 과 같은 구조)",
     "orig_reason": "8월 전문평가 FINAL_INCLUDE(medium) — 불가리아 거주지 Lden × 설문 신체활동. 추가 전문평가(nft_02) 중 "
                    "R1 배제 사례(779·880·1141)와 구조가 같음을 발견해 재판정"},
]

# Crossref 에 저자 메타데이터가 없는 새 포함 연구 — 원문 첫 쪽 저자 줄에서 확인한 값(추측 아님)
AUTHOR_OVERRIDES = {
    "1191": {"authors": ["Sun", "Cheng", "Tian", "Zhou"], "issued": 2026,
             "via": "article PDF p.1 byline (Qing Sun, Qianni Cheng, Jianlin Tian, Yanan Zhou); "
                    "Crossref author 없음 · OpenAlex authorship(X. P. Liu, Z. H. Duan)는 원문과 불일치"},
}

# 미확보 사유 표기 (retrieval_status_all.csv · fig_data.py)
NR_LABEL = {"closed": "No institution access", "nodoi": "No DOI (not located)",
            "oa_fail": "Open-access copy not downloadable", "ppv": "Pay-per-view only",
            "blocked": "Publisher blocked download"}

DOWNSTREAM = ["merge_corpus_v3.py", "finalize_oa_supp_branch.py", "merge_quality_v2.py", "rebuild_table1.py",
              "rebuild_evidence_map.py", "make_fig7_geo_time.py", "fig_data.py", "make_figures.py",
              "make_fig1_prisma_v2.py", "make_fig8_quality.py", "build_tables_en.py", "build_appendix.py",
              "manuscript_facts.py"]
VERIFY = ["verify_manuscript.py", "verify_consistency_v2.py"]   # 원고가 옛 수치면 실패가 정상 — 멈추지 않는다
MA_CLUSTERS = ("walking_speed", "staying", "social", "correlation")


# ═══════════════════ 입출력 ═══════════════════
class CsvFile:
    """기존 파일의 BOM·열 순서를 보존해 다시 쓴다."""

    def __init__(self, path, fields=None, bom=None):
        self.path = path
        self.exists = os.path.exists(path)
        if self.exists:
            raw = open(path, "rb").read(3)
            self.bom = raw == b"\xef\xbb\xbf"
            with open(path, encoding="utf-8-sig", newline="") as f:
                rd = csv.DictReader(f)
                self.rows = list(rd)
                self.fields = list(rd.fieldnames or [])
        else:
            self.bom, self.rows, self.fields = bool(bom), [], list(fields or [])
        self.dirty = False

    def write(self):
        enc = "utf-8-sig" if self.bom else "utf-8"
        os.makedirs(os.path.dirname(self.path), exist_ok=True)
        with open(self.path, "w", newline="", encoding=enc) as f:
            w = csv.DictWriter(f, fieldnames=self.fields, extrasaction="ignore")
            w.writeheader()
            w.writerows(self.rows)


def rd(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        return list(csv.DictReader(f))


def rel(p):
    return os.path.relpath(p, BASE).replace("\\", "/")


def num(uid):
    return int(re.sub(r"\D", "", uid))


def content_hash(rec):
    """배치명 등 비본질 열을 뺀 판정·추출·품질 내용의 해시(원장 대조용)."""
    def strip(r):
        return {k: (v or "").strip() for k, v in sorted(r.items()) if k not in ("batch",)}
    blob = {"verdict": strip({k: v for k, v in rec["verdict"].items()
                              if k not in ("first_verdict", "recheck_verdict", "decision")}),
            "extract": strip(rec["extract"]) if rec["extract"] else None,
            "mmat": strip(rec["mmat"]) if rec["mmat"] else None,
            "detail": sorted((strip(d) for d in rec["detail"]), key=lambda d: d.get("item_no", ""))}
    return hashlib.sha1(json.dumps(blob, ensure_ascii=False, sort_keys=True).encode("utf-8")).hexdigest()[:16]


# ═══════════════════ 1. 입력 세트 읽기·검사 ═══════════════════
def load_sets(set_dirs, errors):
    recs = {}
    for sd in set_dirs:
        sd = os.path.normpath(sd if os.path.isabs(sd) else os.path.join(BASE, sd))
        key = re.sub(r"^newft_final", "newft", os.path.basename(sd)) or "newft"
        need = {n: os.path.join(sd, f"{n}.csv") for n in ("verdict", "extract", "mmat", "detail")}
        miss = [p for p in need.values() if not os.path.exists(p)]
        if miss:
            errors.append(f"[{key}] 입력 파일 없음: {[rel(p) for p in miss]}")
            continue
        ver, ext, mm, det = (rd(need[n]) for n in ("verdict", "extract", "mmat", "detail"))
        es = rd(os.path.join(sd, "es.csv")) if os.path.exists(os.path.join(sd, "es.csv")) else []
        by_ext = defaultdict(list); by_mm = defaultdict(list); by_det = defaultdict(list); by_es = defaultdict(list)
        for r in ext: by_ext[r["id"].strip()].append(r)
        for r in mm: by_mm[r["id"].strip()].append(r)
        for r in det: by_det[r["id"].strip()].append(r)
        for r in es: by_es[r["id"].strip()].append(r)
        for v in ver:
            uid, br, vd = v["id"].strip(), v["branch"].strip(), v["verdict"].strip()
            tag = f"[{key}] {uid}"
            if uid in recs:
                errors.append(f"{tag}: 세트 사이 id 중복({recs[uid]['set']})"); continue
            if br not in ID_RE or not ID_RE[br].match(uid):
                errors.append(f"{tag}: branch={br!r} 와 id 형식 불일치"); continue
            if vd not in VERDICTS:
                errors.append(f"{tag}: verdict={vd!r}"); continue
            code, rule = v.get("reason_code", "").strip(), v.get("boundary_rule", "").strip()
            if vd == "FINAL_EXCLUDE" and code not in X_CATEGORY:
                errors.append(f"{tag}: 배제인데 reason_code={code!r}")
            if vd != "FINAL_EXCLUDE" and code:
                errors.append(f"{tag}: 포함·SENS 인데 reason_code={code!r}")
            if rule not in RULES:
                errors.append(f"{tag}: boundary_rule={rule!r}")
            if (v.get("confidence") or "").strip() not in ("high", "medium", "low"):
                errors.append(f"{tag}: confidence={v.get('confidence')!r}")
            rec = {"uid": uid, "branch": br, "set": key, "set_dir": rel(sd), "verdict": v, "extract": None, "mmat": None,
                   "detail": by_det.get(uid, []), "es": by_es.get(uid, [])}
            if vd in KEEP:
                if len(by_ext[uid]) != 1 or len(by_mm[uid]) != 1:
                    errors.append(f"{tag}: 포함·SENS 인데 extract {len(by_ext[uid])}행 · mmat {len(by_mm[uid])}행")
                else:
                    rec["extract"], rec["mmat"] = by_ext[uid][0], by_mm[uid][0]
                    if rec["extract"].get("final_verdict", vd) != vd:
                        errors.append(f"{tag}: extract.final_verdict 와 verdict 불일치")
                    m = rec["mmat"]
                    qs = [(m.get(q) or "").strip().upper() for q in Q_ITEMS]
                    if any(q not in ("Y", "N", "CT") for q in qs):
                        errors.append(f"{tag}: MMAT 문항 값 {qs}")
                    ny = sum(q == "Y" for q in qs)
                    tier = "high" if ny >= 4 else ("moderate" if ny == 3 else "low")
                    if str(ny) != (m.get("n_yes") or "").strip() or tier != (m.get("quality_tier") or "").strip():
                        errors.append(f"{tag}: n_yes/quality_tier 불일치 (재계산 {ny}/{tier})")
                    cat = (m.get("mmat_category") or "").strip().lower()
                    items = sorted(d.get("item_no", "").strip() for d in rec["detail"])
                    want = [f"{CAT_PREFIX.get(cat, '?')}.{i}" for i in range(1, 6)]
                    if items != want:
                        errors.append(f"{tag}: detail item_no {items} ≠ 범주 {cat!r} 의 {want}")
                    if any((d.get("verdict") or "").strip().upper() not in ("Y", "N", "CT") for d in rec["detail"]):
                        errors.append(f"{tag}: detail verdict 값 오류")
            elif by_ext.get(uid) or by_mm.get(uid) or by_det.get(uid):
                errors.append(f"{tag}: 배제인데 extract/mmat/detail 행이 있음")
            recs[uid] = rec
        orphans = (set(by_ext) | set(by_mm) | set(by_det)) - {v["id"].strip() for v in ver}
        if orphans:
            errors.append(f"[{key}] verdict 에 없는 id 가 extract/mmat/detail 에 있음: {sorted(orphans)}")
    return recs


def load_manifest(patterns):
    meta = {}
    files = []
    for p in patterns:
        files += sorted(glob.glob(p if os.path.isabs(p) else os.path.join(BASE, p)))
    for f in files:
        for r in rd(f):
            meta[r["id"].strip()] = r
    return meta, files


def acquisition_round():
    """id → (날짜 YYYYMMDD, 근거) : 무료 원문 자동 확보(db_oa_retrieval_*.csv OK_PDF/OK_XML)를 먼저,
    그다음 수동 편입 기록(manual_pdf_ingest_*.csv)을 본다(자동 확보분도 편입 기록에 'auto' 로 다시 나온다)."""
    got = {}
    for f in sorted(glob.glob(os.path.join(FT, "db_oa_retrieval_*.csv"))):
        stamp = re.search(r"(\d{8})_(\d{6})", os.path.basename(f))
        for r in rd(f):
            if r["status"].startswith("OK") and r["no"] not in got:
                got[r["no"]] = (stamp.group(1), f"{os.path.basename(f)}:{r['status']}")
    for f in sorted(glob.glob(os.path.join(FT, "manual_pdf_ingest_*.csv"))):
        stamp = re.search(r"(\d{8})_(\d{6})", os.path.basename(f))
        for r in rd(f):
            if r["id"] and r["action"] in ("moved", "copied", "already_in_library", "name_exists") and r["id"] not in got:
                got[r["id"]] = (stamp.group(1), f"{os.path.basename(f)}:{r['origin']}")
    return got


def load_batches():
    """id → 평가 배치명(nft_XX) — newft_batches_*.json."""
    out = {}
    for f in sorted(glob.glob(os.path.join(FT, "newft_batches_*.json"))):
        for b in json.load(open(f, encoding="utf-8")):
            for i in b.get("ids", []):
                out[str(i)] = b["batch"]
    return out


# ═══════════════════ 2. 계획 수립(메모리 안에서만) ═══════════════════
def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("sets", nargs="*", help="판정 세트 폴더 (예: fulltext/newft_final)")
    ap.add_argument("--manifest", nargs="*", default=["fulltext/newft_manifest_*.csv"])
    ap.add_argument("--changes", help="추가 판정 변경 CSV (uid,branch,old,new,rule,reason_code,reason,orig_reason)")
    ap.add_argument("--apply", action="store_true", help="실제로 쓴다(없으면 검사·계획만)")
    ap.add_argument("--fetch-authors", action="store_true", help="새 포함 연구 저자 표기를 Crossref 로 조회해 부록 B 조회표에 추가")
    ap.add_argument("--downstream", action="store_true", help="지도 §3 순서로 다운스트림 재생성")
    ap.add_argument("--log-dir", help="다운스트림 로그 저장 폴더")
    ap.add_argument("--ma-candidates", help="기존 4 클러스터 효과크기 후보를 이 CSV 로 저장")
    a = ap.parse_args()

    if a.sets:
        ok = integrate(a)
        if not ok:
            sys.exit(2)
    if a.downstream:
        if a.sets and not a.apply:
            print("\n⚠️ --downstream 은 --apply 와 함께(또는 세트 없이 단독으로) 써야 한다 — 건너뜀")
        else:
            sys.exit(run_downstream(a.log_dir))


def integrate(a):
    errors, notes = [], []
    recs = load_sets(a.sets, errors)
    meta, man_files = load_manifest(a.manifest)
    changes = list(CORPUS_CHANGES)
    if a.changes:
        changes += rd(a.changes if os.path.isabs(a.changes) else os.path.join(BASE, a.changes))
    acq = acquisition_round()
    batches = load_batches()

    for uid, r in recs.items():
        if uid not in meta:
            errors.append(f"{uid}: manifest 에 서지 없음 ({[rel(f) for f in man_files]})")
        elif meta[uid].get("branch", r["branch"]) != r["branch"]:
            errors.append(f"{uid}: manifest branch={meta[uid].get('branch')} ≠ verdict branch={r['branch']}")
        elif r["extract"]:
            for k in ("year", "title"):
                if (r["extract"].get(k) or "").strip() and r["extract"][k].strip() != meta[uid][k].strip():
                    notes.append(f"{uid}: extract.{k} ≠ manifest.{k} — manifest 값을 서지로 쓴다")
        if uid not in acq:
            notes.append(f"{uid}: 확보 기록(manual_pdf_ingest·db_oa_retrieval)에서 찾지 못함")

    # ── 대상 파일 ──
    F = {
        "ftv": CsvFile(os.path.join(FT, "ft_verdicts_v2.csv")),
        "fte": CsvFile(os.path.join(FT, "ft_extraction_v2.csv")),
        "qall": CsvFile(os.path.join(FT, "quality_all.csv")),
        "qdet": CsvFile(os.path.join(FT, "quality_detail_all.csv")),
        "ctv": CsvFile(os.path.join(FT, "ct_verdicts_final.csv")),
        "cte": CsvFile(os.path.join(FT, "ct_extraction_final.csv")),
        "ctr": CsvFile(os.path.join(FT, "ct_retrieval_final.csv")),
        "oar": CsvFile(os.path.join(FT, "oa_supp_retrieval.csv")),
        "oasq": CsvFile(os.path.join(QA, "qa_oas_mmat.csv")),
        "oasd": CsvFile(os.path.join(QA, "qa_oas_detail.csv")),
        "audit": CsvFile(os.path.join(FT, "ruling_audit.csv")),
        "ledger": CsvFile(os.path.join(FT, "newft_integration_ledger.csv"),
                          fields=["uid", "branch", "kind", "verdict", "reason_code", "boundary_rule",
                                  "source_set", "content_hash", "integrated_at"], bom=True),
    }
    set_files = {}   # 세트별 새 파일 (CT 품질 · OAS 판정/추출)

    def set_file(kind, key):
        tmpl = {"ctq": (QA, "qa_ctfull_{k}_mmat.csv", "qa_ctfull_05_mmat.csv"),
                "ctd": (QA, "qa_ctfull_{k}_detail.csv", "qa_ctfull_05_detail.csv"),
                "oav": (OASFT, "oasft_{k}_verdict.csv", "oasft_01_verdict.csv"),
                "oae": (OASFT, "oasft_{k}_extract.csv", "oasft_01_extract.csv")}[kind]
        path = os.path.join(tmpl[0], tmpl[1].format(k=key))
        if path not in set_files:
            sib = CsvFile(os.path.join(tmpl[0], tmpl[2]))          # 같은 종류 기존 파일의 열·BOM 관례
            set_files[path] = CsvFile(path, fields=sib.fields, bom=sib.bom)
        return set_files[path]

    # 기존 id 집합 (경로별)
    oas_verdict_files = {p: CsvFile(p) for p in sorted(glob.glob(os.path.join(OASFT, "oasft_*_verdict.csv")))}
    ctq_files = {p: CsvFile(p) for p in sorted(glob.glob(os.path.join(QA, "qa_ctfull_*_mmat.csv")))
                 + sorted(glob.glob(os.path.join(QA, "qa_ct_[ABC]_mmat.csv")))}
    existing = {
        "DB": {r["no"].strip() for r in F["ftv"].rows},
        "CT": {f"CT{int(r['rec']):04d}" for r in F["ctv"].rows},
        "OAS": {f"OAS{int(r['sid']):04d}" for cf in oas_verdict_files.values() for r in cf.rows},
    }
    existing_aux = {  # 판정 파일 밖(추출·품질)에 id 가 먼저 들어가 있는 불일치도 잡는다
        "DB": {r["no"].strip() for r in F["fte"].rows} | {r["no"].strip() for r in F["qall"].rows},
        "CT": {"CT%04d" % num(r["no"]) for r in F["cte"].rows}
              | {"CT%04d" % num(r["no"]) for cf in ctq_files.values() for r in cf.rows},
        "OAS": {r["no"].strip() for r in F["oasq"].rows},
    }
    ledger = {r["uid"]: r for r in F["ledger"].rows}

    todo, skipped = [], []
    for uid, r in sorted(recs.items(), key=lambda kv: (kv[1]["branch"], num(kv[0]))):
        h = content_hash(r)
        r["hash"] = h
        led = ledger.get(uid)
        if led:
            if led["kind"] != "new-study" or led["content_hash"] != h:
                errors.append(f"{uid}: 원장에 {led['kind']}·해시 {led['content_hash']} 로 이미 반영됨 — 이번 입력(해시 {h})과 "
                              f"다르다. 판정이 바뀌었으면 수동으로 되돌린 뒤 다시 실행")
            elif uid not in existing[r["branch"]]:
                errors.append(f"{uid}: 원장에는 반영 기록이 있는데 판정 파일에 행이 없다 — 정본 불일치")
            else:
                skipped.append(uid)
            continue
        if uid in existing[r["branch"]] or uid in existing_aux[r["branch"]]:
            errors.append(f"{uid}: 원장에 없는 id 가 기존 {r['branch']} 정본에 이미 있다 — 중단(중복 반영 방지)")
            continue
        todo.append(r)

    ch_todo = []
    for c in changes:
        uid = c["uid"].strip()
        led = ledger.get(uid)
        cur = next((x for x in F["ftv"].rows if x["no"] == uid), None) if c["branch"] == "DB" else None
        if c["branch"] != "DB":
            errors.append(f"{uid}: 판정 변경은 현재 DB 경로만 지원"); continue
        if c["new"] not in VERDICTS or c["old"] not in VERDICTS:
            errors.append(f"{uid}: 판정 변경 값 오류 {c['old']}→{c['new']}"); continue
        if c["old"] == "FINAL_EXCLUDE" and c["new"] in KEEP:
            errors.append(f"{uid}: 배제→포함 변경은 추출·품질 자료가 필요 — 새 판정 세트로 넣을 것"); continue
        if led:
            if led["kind"] != "ruling-change" or led["verdict"] != c["new"]:
                errors.append(f"{uid}: 원장 기록({led['kind']} {led['verdict']})과 판정 변경({c['new']})이 다르다")
            elif not cur or cur["final_verdict"] != c["new"]:
                errors.append(f"{uid}: 원장에는 {c['new']} 로 반영됐는데 ft_verdicts_v2 는 {cur and cur['final_verdict']}")
            else:
                skipped.append(uid)
            continue
        if not cur:
            errors.append(f"{uid}: ft_verdicts_v2 에 없음"); continue
        if cur["final_verdict"] != c["old"]:
            errors.append(f"{uid}: 현재 판정 {cur['final_verdict']} ≠ 변경 전 기대값 {c['old']} — 중단"); continue
        if uid in recs:
            errors.append(f"{uid}: 판정 변경 대상이 새 판정 세트에도 있다"); continue
        ch_todo.append(c)

    print(f"=== 통합 점검 ({NOW}) ===")
    print(f"입력 세트 {len(a.sets)}개 · 판정 {len(recs)}건 {dict(Counter((r['branch'], r['verdict']['verdict']) for r in recs.values()))}")
    print(f"반영 예정 {len(todo)}건 · 판정 변경 {len(ch_todo)}건 · 이미 반영돼 건너뜀 {len(skipped)}건")
    for n in notes:
        print("  ·", n)
    if errors:
        print("\n❌ 중단 — 아무것도 쓰지 않았다:")
        for e in errors:
            print("  -", e)
        return False

    # ── 3. 정본에 행 추가 (메모리) ──
    def bib(uid):
        m = meta[uid]
        return {"year": m["year"].strip(), "journal": m["journal"].strip(), "title": m["title"].strip()}

    def ext_row(r, no):
        e = {c: (r["extract"].get(c) or "").strip() for c in EXT_COLS}
        e.update(bib(r["uid"]), no=no, final_verdict=r["verdict"]["verdict"])
        return e

    def mmat_row(r, no, title_case):
        m = r["mmat"]
        cat = m["mmat_category"].strip()
        cat = cat[0].upper() + cat[1:].lower() if title_case else cat.lower()
        cat = cat.replace("rct", "RCT")
        out = {"no": no, "mmat_category": cat, "S1": m["S1"].strip().upper(), "S2": m["S2"].strip().upper(),
               **{q: m[q].strip().upper() for q in Q_ITEMS},
               "n_yes": m["n_yes"].strip(), "quality_tier": m["quality_tier"].strip(), "note": (m.get("note") or "").strip()}
        return out

    date_tag = lambda uid: acq.get(uid, (datetime.now(KST).strftime("%Y%m%d"), "기록 없음"))
    ledger_rows = []
    for r in todo:
        uid, br, v = r["uid"], r["branch"], r["verdict"]
        vd, code, rule = v["verdict"].strip(), v["reason_code"].strip(), v["boundary_rule"].strip()
        d8, src = date_tag(uid)
        d_iso = f"{d8[:4]}-{d8[4:6]}-{d8[6:]}"
        what = ("포함" if vd == "FINAL_INCLUDE" else "민감도 전용" if vd == "SENS_ONLY" else f"배제 {code} {X_KO[code]}")
        reason = (f"추가 전문평가({d_iso} 확보): {what}" + (f" · 경계규칙 {rule}" if rule else "")
                  + f" — 근거 {r['set_dir']}/verdict.csv")
        if br == "DB":
            F["ftv"].rows.append({"no": uid, "final_verdict": vd, "confidence": v["confidence"].strip(), "reason": reason,
                                  "screening_verdict": meta[uid]["screening"].strip(), **bib(uid)})
            if vd in KEEP:
                F["fte"].rows.append(ext_row(r, uid))
                F["qall"].rows.append({**mmat_row(r, uid, True), **bib(uid)})
                for d in sorted(r["detail"], key=lambda d: d["item_no"]):
                    F["qdet"].rows.append({"no": uid, "item": f"{d['item_no'].strip()} {d['item'].strip()}",
                                           "item_no": d["item_no"].strip(), "verdict": d["verdict"].strip().upper(),
                                           "rationale": d["rationale"].strip()})
        elif br == "CT":
            rec = str(num(uid))
            batch = (r["extract"] or {}).get("batch") or batches.get(uid) or r["set"]
            F["ctv"].rows.append({"rec": rec, "verdict": vd, "reason_code": code, "rationale": v["rationale"].strip(),
                                  "boundary_rule": rule, "batch": batch, "screen1": meta[uid]["screening"].strip(),
                                  "doi": meta[uid]["doi"].strip(), **bib(uid)})
            if vd in KEEP:
                F["cte"].rows.append(ext_row(r, rec))
                set_file("ctq", r["set"]).rows.append(mmat_row(r, uid, False))
                for d in sorted(r["detail"], key=lambda d: d["item_no"]):
                    set_file("ctd", r["set"]).rows.append({"no": uid, "item": d["item"].strip(), "item_no": d["item_no"].strip(),
                                                          "verdict": d["verdict"].strip().upper(),
                                                          "rationale": d["rationale"].strip()})
            hit = [x for x in F["ctr"].rows if f"CT{int(x['rec']):04d}" == uid]
            if not hit:
                print(f"❌ {uid}: ct_retrieval_final.csv(전문 확보 대상)에 없음 — 중단"); return False
            if hit[0]["state"] == "not-retrieved":
                hit[0]["reason"] = f"8월 미확보({hit[0]['reason']}) → {d_iso} 추가 확보 [{src}]"
                hit[0]["state"] = f"retrieved-{d8}"
        else:  # OAS
            sid = str(num(uid))
            set_file("oav", r["set"]).rows.append({"sid": sid, "verdict": vd, "reason_code": code,
                                                   "rationale": v["rationale"].strip(), "boundary_rule": rule})
            if vd in KEEP:
                set_file("oae", r["set"]).rows.append(ext_row(r, uid))
                F["oasq"].rows.append(mmat_row(r, uid, False))
                for d in sorted(r["detail"], key=lambda d: d["item_no"]):
                    F["oasd"].rows.append({"no": uid, "item": d["item"].strip(), "item_no": d["item_no"].strip(),
                                           "verdict": d["verdict"].strip().upper(), "rationale": d["rationale"].strip()})
            hit = [x for x in F["oar"].rows if f"OAS{int(x['sid']):04d}" == uid]
            if not hit:
                print(f"❌ {uid}: oa_supp_retrieval.csv(전문 확보 대상)에 없음 — 중단"); return False
            if not (hit[0]["result"] in ("ok", "already") or hit[0]["result"].startswith("ok-")):
                hit[0]["reason"] = f"8월 미확보({hit[0]['result']}) → {d_iso} 추가 확보 [{src}]"
                hit[0]["result"] = f"ok-{d8}"
        ledger_rows.append({"uid": uid, "branch": br, "kind": "new-study", "verdict": vd, "reason_code": code,
                            "boundary_rule": rule, "source_set": r["set"], "content_hash": r["hash"], "integrated_at": NOW})

    for c in ch_todo:
        uid = c["uid"]
        cur = next(x for x in F["ftv"].rows if x["no"] == uid)
        cur["final_verdict"], cur["reason"] = c["new"], c["reason"]
        if c["new"] not in KEEP:
            for k in ("fte", "qall", "qdet"):
                n0 = len(F[k].rows)
                F[k].rows = [x for x in F[k].rows if x["no"].strip() != uid]
                print(f"  판정 변경 {uid}: {rel(F[k].path)} 에서 {n0 - len(F[k].rows)}행 제거")
        else:
            for x in F["fte"].rows:
                if x["no"] == uid:
                    x["final_verdict"] = c["new"]
        F["audit"].rows.append({"no": uid, "rule": c["rule"], "old": c["old"], "new": c["new"],
                                "orig_reason": c["orig_reason"], "title": cur["title"][:70]})
        ledger_rows.append({"uid": uid, "branch": "DB", "kind": "ruling-change", "verdict": c["new"],
                            "reason_code": c["reason_code"], "boundary_rule": c["rule"], "source_set": "CORPUS_CHANGES",
                            "content_hash": "", "integrated_at": NOW})

    # 정렬: 기존 파일 관례(id 오름차순)
    F["ftv"].rows.sort(key=lambda x: int(x["no"]))
    F["fte"].rows.sort(key=lambda x: int(x["no"]))
    F["qall"].rows.sort(key=lambda x: int(x["no"]))
    F["qdet"].rows.sort(key=lambda x: int(x["no"]))            # 안정 정렬 — 연구 안 문항 순서 유지
    F["ctv"].rows.sort(key=lambda x: int(x["rec"]))
    F["cte"].rows.sort(key=lambda x: int(x["no"]))
    F["ledger"].rows += ledger_rows

    # ── 4. 파생 표 재생성 ──
    led_all = {x["uid"]: x for x in F["ledger"].rows}
    excl_rows, bad = build_exclusion_reasons(F, changes, led_all, set_files)
    ret_rows, bad2 = build_retrieval_status(F, acq, set_files)
    if bad or bad2:
        print("\n❌ 파생 표 검사 실패 — 아무것도 쓰지 않았다:")
        for e in bad + bad2:
            print("  -", e)
        return False

    # ── 5. 반영 후 예상 수치 ──
    report_expected(F, excl_rows, ret_rows, set_files)
    if a.ma_candidates:
        write_ma_candidates(recs, meta, a.ma_candidates)

    if not a.apply:
        print("\n(검사·계획만 수행 — 쓰려면 --apply)")
        return True

    written = []
    for k, cf in F.items():
        cf.write(); written.append(rel(cf.path))
    for cf in set_files.values():
        if cf.rows:
            cf.write(); written.append(rel(cf.path))
    write_plain(os.path.join(FT, "ft_exclusion_reasons.csv"),
                ["uid", "branch", "reason_code", "category", "basis"], excl_rows)
    write_plain(os.path.join(FT, "retrieval_status_all.csv"),
                ["uid", "branch", "screening", "status", "round", "nr_category", "evidence"], ret_rows)
    written += ["fulltext/ft_exclusion_reasons.csv", "fulltext/retrieval_status_all.csv"]
    print("\n[저장] " + " · ".join(written))
    if a.fetch_authors:
        fetch_authors(F, meta)
    return True


def write_plain(path, fields, rows):
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.DictWriter(f, fieldnames=fields)
        w.writeheader(); w.writerows(rows)


# ═══════════════════ 파생 표 ═══════════════════
def oas_verdict_rows(set_files):
    """디스크의 oasft_*_verdict.csv + 이번 실행에서 메모리에 만든 세트 파일(아직 안 쓴 것 포함)."""
    paths = set(glob.glob(os.path.join(OASFT, "oasft_*_verdict.csv")))
    paths |= {p for p in set_files if p.endswith("_verdict.csv")}
    rows = []
    for p in sorted(paths):
        rows += set_files[p].rows if p in set_files else CsvFile(p).rows
    return rows


def build_exclusion_reasons(F, changes, ledger, set_files):
    out, bad = [], []
    change_code = {c["uid"]: c for c in changes}
    for r in F["ftv"].rows:
        if r["final_verdict"] != "FINAL_EXCLUDE":
            continue
        uid = r["no"]
        if uid in change_code and ledger.get(uid, {}).get("kind") == "ruling-change":
            code, basis = change_code[uid]["reason_code"], f"8월 코퍼스 판정 변경 {change_code[uid]['rule']}"
        elif uid in AUG_DB_EXCLUSION:
            code, basis = AUG_DB_EXCLUSION[uid], "8월 수기 집계 재배정(ft_verdicts_v2.reason)"
        elif ledger.get(uid, {}).get("kind") == "new-study":
            code, basis = ledger[uid]["reason_code"], f"추가 전문평가 {ledger[uid]['source_set']}"
        else:
            bad.append(f"DB {uid}: 전문 배제 사유 코드를 찾을 수 없음"); continue
        out.append({"uid": uid, "branch": "DB", "reason_code": code,
                    "category": SPECIAL_CATEGORY.get(code) or X_CATEGORY.get(code, "?"), "basis": basis})
    for r in F["ctv"].rows:
        if r["verdict"] == "FINAL_EXCLUDE":
            code = r["reason_code"].strip()
            if code not in X_CATEGORY:
                bad.append(f"CT{int(r['rec']):04d}: reason_code={code!r}"); continue
            out.append({"uid": f"CT{int(r['rec']):04d}", "branch": "CT", "reason_code": code,
                        "category": X_CATEGORY_BRANCH.get(("CT", code), X_CATEGORY[code]),
                        "basis": f"ct_verdicts_final.csv ({r['batch']})"})
    seen = set()
    for r in oas_verdict_rows(set_files):
        uid = f"OAS{int(r['sid']):04d}"
        if uid in seen:
            bad.append(f"{uid}: oasft 판정 파일에 중복"); continue
        seen.add(uid)
        if r["verdict"] == "FINAL_EXCLUDE":
            code = r["reason_code"].strip()
            if code not in X_CATEGORY:
                bad.append(f"{uid}: reason_code={code!r}"); continue
            out.append({"uid": uid, "branch": "OAS", "reason_code": code,
                        "category": X_CATEGORY_BRANCH.get(("OAS", code), X_CATEGORY[code]),
                        "basis": "oasft verdict"})
    return out, bad


def build_retrieval_status(F, acq, set_files):
    out, bad = [], []
    # DB — 확보 대상 189 = fulltext_status_20260803.csv, 평가 = ft_verdicts_v2.csv
    status = rd(os.path.join(FT, "fulltext_status_20260803.csv"))
    oa_files = sorted(glob.glob(os.path.join(FT, "db_oa_retrieval_*.csv")))
    oa = {r["no"]: r for r in rd(oa_files[-1])} if oa_files else {}
    assessed = {r["no"] for r in F["ftv"].rows}
    for r in status:
        uid = r["no"]
        if uid in assessed:
            rnd, ev = ("2026-08", "fulltext_status_20260803.csv pdf=Y") if r["pdf"] == "Y" else \
                (acq[uid][0] if uid in acq else "?", acq.get(uid, ("", "기록 없음"))[1])
            out.append({"uid": uid, "branch": "DB", "screening": r["verdict"], "status": "retrieved",
                        "round": rnd, "nr_category": "", "evidence": ev})
            if r["pdf"] != "Y" and uid not in acq:
                bad.append(f"DB {uid}: 평가됐는데 확보 기록이 없음")
        else:
            if r["pdf"] == "Y":
                bad.append(f"DB {uid}: 8월 확보(pdf=Y)인데 전문 판정이 없음")
            o = oa.get(uid, {})
            if not (o.get("doi") or r.get("doi") or "").strip():
                cat = NR_LABEL["nodoi"]
            elif (o.get("oa_status") or "").strip() in ("", "closed"):
                cat = NR_LABEL["closed"]
            else:
                cat = NR_LABEL["oa_fail"]
            out.append({"uid": uid, "branch": "DB", "screening": r["verdict"], "status": "not-retrieved",
                        "round": "", "nr_category": cat,
                        "evidence": f"{os.path.basename(oa_files[-1]) if oa_files else ''}: oa_status={o.get('oa_status','')} "
                                    f"tried={o.get('tried','')}; 수동 편입 기록 없음"})
    extra = assessed - {r["no"] for r in status}
    if extra:
        bad.append(f"DB 전문 판정이 확보 대상(189) 밖에 있음: {sorted(extra)}")
    # CT
    ct_assessed = {f"CT{int(r['rec']):04d}" for r in F["ctv"].rows}
    for r in F["ctr"].rows:
        uid = f"CT{int(r['rec']):04d}"
        ok = r["state"].startswith("retrieved")
        if ok != (uid in ct_assessed):
            bad.append(f"{uid}: 확보 상태({r['state']})와 전문 판정 유무({uid in ct_assessed}) 불일치")
        cat = "" if ok else (NR_LABEL["ppv"] if "pay-per-view" in r["reason"].lower() else NR_LABEL["closed"])
        out.append({"uid": uid, "branch": "CT", "screening": r["screen"], "status": "retrieved" if ok else "not-retrieved",
                    "round": ("2026-08" if r["state"] == "retrieved" else r["state"].replace("retrieved-", "")) if ok else "",
                    "nr_category": cat, "evidence": f"ct_retrieval_final.csv: {r['state']} {r['reason']}"})
    extra = ct_assessed - {f"CT{int(r['rec']):04d}" for r in F["ctr"].rows}
    if extra:
        bad.append(f"CT 전문 판정이 확보 대상 밖: {sorted(extra)}")
    # OAS
    oas_assessed = {f"OAS{int(r['sid']):04d}" for r in oas_verdict_rows(set_files)}
    for r in F["oar"].rows:
        uid = f"OAS{int(r['sid']):04d}"
        res = r["result"]
        if res == "prescreen-exclude":
            st, cat = "prescreen-exclude", ""
        elif res in ("ok", "already") or res.startswith("ok-"):
            st, cat = "retrieved", ""
        else:
            st = "not-retrieved"
            cat = NR_LABEL["blocked"] if res.startswith("http") else NR_LABEL["closed"]
        if (st == "retrieved") != (uid in oas_assessed):
            bad.append(f"{uid}: 확보 상태({res})와 전문 판정 유무({uid in oas_assessed}) 불일치")
        out.append({"uid": uid, "branch": "OAS", "screening": r["verdict"], "status": st,
                    "round": ("2026-08" if res in ("ok", "already") else res.replace("ok-", "")) if st == "retrieved" else "",
                    "nr_category": cat, "evidence": f"oa_supp_retrieval.csv: {res} {r['reason']}"})
    return out, bad


def report_expected(F, excl_rows, ret_rows, set_files):
    tot = Counter()
    print("\n=== 반영 후 정본 판정 (예상) ===")
    for b, cnt in (("DB", Counter(r["final_verdict"] for r in F["ftv"].rows)),
                   ("CT", Counter(r["verdict"] for r in F["ctv"].rows)),
                   ("OAS", Counter(r["verdict"] for r in oas_verdict_rows(set_files)))):
        tot.update(cnt)
        print(f"  {b:3s} 평가 {sum(cnt.values())} · 포함 {cnt['FINAL_INCLUDE']} · SENS {cnt['SENS_ONLY']} · 배제 {cnt['FINAL_EXCLUDE']}")
    print(f"  합계 포함 {tot['FINAL_INCLUDE']} · SENS {tot['SENS_ONLY']} · 분석 대상 {tot['FINAL_INCLUDE'] + tot['SENS_ONLY']}")
    by = defaultdict(Counter)
    for r in excl_rows:
        by[r["branch"]][r["category"]] += 1
    for b in ("DB", "CT", "OAS"):
        print(f"  전문 배제 사유 {b}: {dict(by[b].most_common())} (합 {sum(by[b].values())})")
    st = defaultdict(Counter)
    nr = defaultdict(Counter)
    for r in ret_rows:
        st[r["branch"]][r["status"]] += 1
        if r["nr_category"]:
            nr[r["branch"]][r["nr_category"]] += 1
    for b in ("DB", "CT", "OAS"):
        print(f"  확보 대상 {b}: {sum(st[b].values())} {dict(st[b])} · 미확보 사유 {dict(nr[b])}")


def write_ma_candidates(recs, meta, out):
    rows = []
    for uid, r in recs.items():
        vd = r["verdict"]["verdict"]
        for e in r["es"]:
            if e["cluster"].strip() in MA_CLUSTERS:
                rows.append({"id": uid, "branch": r["branch"], "verdict": vd, "cluster": e["cluster"].strip(),
                             "computable": e["computable"].strip(), "outcome_measure": e["outcome_measure"],
                             "comparison": e["comparison"], "statistic_type": e["statistic_type"],
                             "values_verbatim": e["values_verbatim"], "n_info": e["n_info"], "location": e["location"]})
    rows.sort(key=lambda x: (x["cluster"], x["branch"], num(x["id"])))
    write_plain(out, list(rows[0].keys()) if rows else ["id"], rows)
    print(f"\n[MA 후보] 기존 4 클러스터 해당 {len(rows)}행 → {out} "
          f"{dict(Counter((x['cluster'], x['computable']) for x in rows))}")


# ═══════════════════ 부록 B 저자 표기 (Crossref) ═══════════════════
def fetch_authors(F, meta):
    lp = os.path.join(FT, "appendix_b_author_lookup.json")
    cp = os.path.join(FT, "appendix_b_crossref_all.json")
    look = json.load(open(lp, encoding="utf-8"))
    cross = json.load(open(cp, encoding="utf-8")) if os.path.exists(cp) else {}
    need = [r["uid"] for r in F["ledger"].rows
            if r["kind"] == "new-study" and r["verdict"] == "FINAL_INCLUDE" and r["uid"] not in look]
    print(f"\n[저자 조회] 새 포함 연구 중 조회표에 없는 {len(need)}편")
    fails = []
    for uid in need:
        doi = (meta.get(uid, {}).get("doi") or "").strip()
        if not doi:
            fails.append(f"{uid}: DOI 없음 — 수기 확인 필요"); continue
        url = "https://api.crossref.org/works/" + urllib.parse.quote(doi, safe="/:;()._-")
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "paper32-systematic-review/1.0"})
            msg = json.load(urllib.request.urlopen(req, timeout=45))["message"]
        except Exception as e:  # noqa: BLE001
            fails.append(f"{uid}: Crossref 조회 실패 {e}"); continue
        fam = [au["family"].strip() for au in msg.get("author", []) if au.get("family")]
        issued = (msg.get("issued", {}).get("date-parts") or [[None]])[0][0]
        via = "api.crossref.org"
        if not fam and uid in AUTHOR_OVERRIDES:
            fam, via = AUTHOR_OVERRIDES[uid]["authors"], AUTHOR_OVERRIDES[uid]["via"]
            issued = issued or AUTHOR_OVERRIDES[uid]["issued"]
        if not fam:
            fails.append(f"{uid}: Crossref 저자 성(family) 없음 — 원문 확인 후 AUTHOR_OVERRIDES 에 추가"); continue
        label = fam[0] if len(fam) == 1 else (f"{fam[0]} & {fam[1]}" if len(fam) == 2 else f"{fam[0]} et al.")
        if not issued:
            fails.append(f"{uid}: Crossref 발행연도 없음"); continue
        look[uid] = {"label": label, "year": issued, "doi": doi,
                     "title": (msg.get("title") or [meta[uid]["title"]])[0][:60], "via": via}
        cross[uid] = {"authors": fam, "issued": issued, "doi": doi}
        yr = meta[uid]["year"].strip()
        print(f"  {uid}: {label} ({issued})" + (f"  ⚠️ 코퍼스 연도 {yr} ≠ Crossref issued {issued}" if yr != str(issued) else ""))
        time.sleep(0.3)
    json.dump(look, open(lp, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    json.dump(cross, open(cp, "w", encoding="utf-8"), ensure_ascii=False, indent=1)
    for f in fails:
        print("  ❌", f)


# ═══════════════════ 다운스트림 ═══════════════════
def run_downstream(log_dir=None):
    env = dict(os.environ, PYTHONIOENCODING="utf-8", MPLBACKEND="Agg")
    if log_dir:
        os.makedirs(log_dir, exist_ok=True)
    print(f"\n=== 다운스트림 재생성 ({NOW}) — 메타분석 스크립트(ma_*.py)는 실행하지 않는다 ===")
    status = 0
    for step in DOWNSTREAM + VERIFY:
        t0 = time.time()
        p = subprocess.run([sys.executable, os.path.join(BASE, step)], cwd=BASE, env=env,
                           capture_output=True, text=True, encoding="utf-8", errors="replace")
        out = p.stdout + p.stderr
        if log_dir:
            open(os.path.join(log_dir, step.replace(".py", ".log")), "w", encoding="utf-8").write(out)
        warns = [ln.strip() for ln in out.splitlines() if "⚠️" in ln or "❌" in ln or "Traceback" in ln]
        mark = "✅" if p.returncode == 0 else ("⚠️" if step in VERIFY else "❌")
        print(f"{mark} {step:30s} exit={p.returncode} ({time.time() - t0:.1f}s)" + (f" · 경고 {len(warns)}" if warns else ""))
        for w in warns[:12]:
            print(f"     {w[:160]}")
        if p.returncode != 0 and step not in VERIFY:
            print("\n".join("     | " + ln for ln in out.strip().splitlines()[-15:]))
            print(f"\n❌ {step} 에서 중단")
            return 1
        if p.returncode != 0:
            status = 3
    return status


if __name__ == "__main__":
    main()
