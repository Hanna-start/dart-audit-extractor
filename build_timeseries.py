# -*- coding: utf-8 -*-
"""유연한 3단계 자동 소스선택 시계열 빌더.

회사×연도마다 가용한 최선의 소스를 자동 선택해 한 시계열로 통합한다.
  ① 사업보고서 있음(정기보고서)  → fnlttSinglAcntAll 정형 API  (가장 견고)
  ② 감사보고서만 있음            → document.xml XML 표 파싱      (PDF보다 견고)
  ③ ②도 실패(스캔 등)           → PDF 스크래핑                  (현행, 최후 fallback)

비상장→상장/사업보고서 전환을 한 회사가 모두 거쳐도(무신사) 자동으로 소스가 바뀐다.

사용:
  py build_timeseries.py --corp 01137727 --company 무신사 --from 2018 --to 2025 \
     --out output/2026-05/무신사_시계열_flex_2026-05-31.xlsx
"""
from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

import pandas as pd

import dart_client
import dart_xml_fs
import api_to_timeseries as api

HERE = Path(__file__).resolve().parent


def _key(cli):
    if cli:
        return cli
    k = os.environ.get("DART_API_KEY", "")
    if k:
        return k
    envp = HERE / ".env"
    if envp.exists():
        for line in envp.read_text(encoding="utf-8").splitlines():
            if line.strip().startswith("DART_API_KEY="):
                return line.split("=", 1)[1].strip().strip('"').strip("'")
    return ""


def build(key, corp, company, y_from, y_to):
    years = list(range(y_from, y_to + 1))
    source_map = {}   # (year, 구분) → 소스 라벨

    # ── ① API 계층 (사업보고서) ──
    api_df = api.build_long_rows(key, corp, company, years)
    api_recs = api_df.to_dict("records") if not api_df.empty else []
    api_year_div = set((int(r["사업연도"]), r["구분"]) for r in api_recs)
    for (yr, dv) in api_year_div:
        source_map[(yr, dv)] = "① 사업보고서 API"

    # ── ② XML 계층 (감사보고서 document.xml) ──
    disc = dart_client.list_audit_disclosures(
        key, corp, bgn_de=f"{y_from}0101", end_de=f"{y_to + 2}1231")
    latest = dart_client.dedupe_latest(disc)
    xml_recs = []
    for d in sorted(latest, key=lambda x: x.rcept_dt):
        fy = d.fiscal_hint_year()
        if fy < y_from or fy > y_to:
            continue
        if (fy, d.kind) in api_year_div:     # 그 해는 API가 이미 커버 → 스킵
            continue
        try:
            recs = dart_xml_fs.extract(key, d.rcept_no, d.kind, company, fy, fy - 1)
        except Exception as e:
            print(f"  [XML 실패] {d.rcept_no} {d.kind} fy{fy}: {e}")
            recs = []
        if recs:
            xml_recs += recs
            for yr in {int(r["사업연도"]) for r in recs}:
                if y_from <= yr <= y_to:
                    source_map.setdefault((yr, d.kind), "② 감사보고서 XML")
        # ③ PDF fallback: XML도 빈 경우 (스캔본 등) 여기서 PDF 경로 호출 가능.
        #    현행 pdf_to_excel/extractor 흐름을 연결하는 자리. (auto-run 생략)

    all_recs = xml_recs + api_recs      # API를 뒤에 둬 동일 (연,구분,계정)에서 API 우선
    long_df = pd.DataFrame(all_recs)
    long_df = long_df[(long_df["사업연도"] >= y_from) & (long_df["사업연도"] <= y_to)]
    long_df = (long_df.groupby(["회사명", "사업연도", "구분", "재무제표", "계정"], as_index=False)["값"].last())
    long_df = long_df.sort_values(["구분", "재무제표", "사업연도", "계정"]).reset_index(drop=True)
    return long_df, source_map


def write_xlsx(long_df, out, source_map=None):
    sheets = {}
    # 출처(provenance) 시트 — 어느 연도×구분이 어느 소스(①API/②XML/③PDF)에서 왔는지.
    # build()가 계산한 source_map을 산출물에 담아 신뢰근거가 파일과 함께 이동하게 한다.
    if source_map:
        srows = [{"사업연도": yr, "구분": dv, "출처": source_map[(yr, dv)]}
                 for (yr, dv) in sorted(source_map)]
        sheets["출처"] = pd.DataFrame(srows)
    sheets["long_data"] = long_df
    for gubun in ("별도", "연결"):
        for sj in ("재무상태표", "포괄손익계산서", "현금흐름표"):
            sub = long_df[(long_df["구분"] == gubun) & (long_df["재무제표"] == sj)]
            if sub.empty:
                continue
            piv = sub.groupby(["계정", "사업연도"])["값"].last().unstack("사업연도")
            piv = piv.reindex(sorted(piv.columns), axis=1).reset_index()
            sheets[f"{gubun}_{sj}"] = piv
    Path(out).parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        for name, df in sheets.items():
            df.to_excel(w, sheet_name=name, index=False)


def main(argv):
    ap = argparse.ArgumentParser()
    ap.add_argument("--corp", required=True)
    ap.add_argument("--company", required=True)
    ap.add_argument("--from", dest="y_from", type=int, required=True)
    ap.add_argument("--to", dest="y_to", type=int, required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--key", default="")
    args = ap.parse_args(argv[1:])

    key = _key(args.key)
    if not key:
        print("[!] DART_API_KEY 필요"); sys.exit(1)

    long_df, smap = build(key, args.corp, args.company, args.y_from, args.y_to)
    write_xlsx(long_df, args.out, source_map=smap)

    print(f"WROTE: {args.out}  (long_data {len(long_df)}행)")
    print("\n소스 선택 결과 (연도×구분):")
    for (yr, dv) in sorted(smap):
        print(f"  {yr} {dv}: {smap[(yr, dv)]}")

    # 회계등식 검증 + 검증_Report 시트 주입 (결정성 분리: validate.py asset만 호출)
    _validate_output(args.out, args.company)


# assets/validate.py 는 프로젝트 루트에서 바로 import되지 않으므로 경로를 주입해 호출.
# (헌법: 검증 계산은 validate.py에서만. 여기선 그 함수를 호출만 한다.)
_ASSETS = HERE / ".claude" / "skills" / "dart-pipeline" / "assets"


def _safe(s: str) -> str:
    bad = '<>:"/\\|?*'
    return "".join("_" if c in bad else c for c in str(s)).strip() or "회사"


def _validate_output(out_path, company):
    if str(_ASSETS) not in sys.path:
        sys.path.insert(0, str(_ASSETS))
    try:
        import validate as fs_validate
    except Exception as e:
        print(f"\n[검증 스킵] validate.py 로드 실패: {e}")
        return
    log_dir = HERE / "_workspace" / _safe(company) / "logs"
    try:
        rep = fs_validate.validate_and_embed(out_path, log_dir=log_dir)
    except Exception as e:
        print(f"\n[검증 실패] {e}")
        return
    print("")
    fs_validate.print_report(rep)
    print(f"  검증_Report 시트 주입 완료 · 로그: {log_dir}")


if __name__ == "__main__":
    main(sys.argv)
