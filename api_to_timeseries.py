# -*- coding: utf-8 -*-
"""사업보고서 정형 API(fnlttSinglAcntAll) → 기존 PDF 시계열에 병합.

무신사처럼 사업보고서 제출 대상으로 전환된 회사의 신규 연도를 PDF 없이 수집한다.
계정 매핑은 account_id(IFRS 표준코드) 기준 — references/assumptions/<회사>/api_account_map.md 와 동기화.

사용:
  py api_to_timeseries.py --corp 01137727 --company 무신사 \
     --base output/2026-05/무신사_시계열_2026-05-30.xlsx \
     --years 2024 2025 --out output/2026-05/무신사_시계열_2026-05-31.xlsx
"""
from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

import pandas as pd
import requests

HERE = Path(__file__).resolve().parent
OPEN = "https://opendart.fss.or.kr/api"

# account_id → (재무제표, 시계열 계정명).  account_id 기반이라 라벨 변형에 강함.
ID_MAP = {
    # 재무상태표
    "ifrs-full_Assets": ("재무상태표", "자산총계"),
    "ifrs-full_Liabilities": ("재무상태표", "부채총계"),
    "ifrs-full_Equity": ("재무상태표", "자본총계"),
    "ifrs-full_CurrentAssets": ("재무상태표", "유동자산"),
    "ifrs-full_NoncurrentAssets": ("재무상태표", "비유동자산"),
    "ifrs-full_CurrentLiabilities": ("재무상태표", "유동부채"),
    "ifrs-full_NoncurrentLiabilities": ("재무상태표", "비유동부채"),
    "ifrs-full_IssuedCapital": ("재무상태표", "자본금"),
    # 포괄손익계산서
    "ifrs-full_Revenue": ("포괄손익계산서", "매출액"),
    "ifrs-full_CostOfSales": ("포괄손익계산서", "매출원가"),
    "ifrs-full_GrossProfit": ("포괄손익계산서", "매출총이익"),
    "dart_OperatingIncomeLoss": ("포괄손익계산서", "영업이익"),
    "dart_TotalSellingGeneralAdministrativeExpenses": ("포괄손익계산서", "판매비와관리비"),
    "ifrs-full_ProfitLossBeforeTax": ("포괄손익계산서", "법인세차감전순이익"),
    "ifrs-full_IncomeTaxExpenseContinuingOperations": ("포괄손익계산서", "법인세비용"),
    "ifrs-full_ProfitLoss": ("포괄손익계산서", "당기순이익"),
    "ifrs-full_ComprehensiveIncome": ("포괄손익계산서", "총포괄이익"),
    # 현금흐름표
    "ifrs-full_CashFlowsFromUsedInOperatingActivities": ("현금흐름표", "영업활동현금흐름"),
    "ifrs-full_CashFlowsFromUsedInInvestingActivities": ("현금흐름표", "투자활동현금흐름"),
    "ifrs-full_CashFlowsFromUsedInFinancingActivities": ("현금흐름표", "재무활동현금흐름"),
    "dart_CashAndCashEquivalentsAtEndOfPeriodCf": ("현금흐름표", "기말현금"),
}
# account_id가 비표준("-표준계정코드 미사용-")일 때 account_nm 포함 매칭 (현금흐름 보조)
NM_FALLBACK = [
    ("영업활동", ("현금흐름표", "영업활동현금흐름")),
    ("투자활동", ("현금흐름표", "투자활동현금흐름")),
    ("재무활동", ("현금흐름표", "재무활동현금흐름")),
    ("기말의 현금", ("현금흐름표", "기말현금")),
]


def _to_int(amt):
    if amt in (None, "", "-"):
        return None
    try:
        return int(str(amt).replace(",", ""))
    except ValueError:
        return None


def fetch_year(key: str, corp: str, year: int, fs_div: str) -> dict:
    """{(재무제표, 계정): 값} 반환. fs_div: OFS=별도, CFS=연결."""
    p = {"crtfc_key": key, "corp_code": corp, "bsns_year": str(year),
         "reprt_code": "11011", "fs_div": fs_div}
    d = requests.get(f"{OPEN}/fnlttSinglAcntAll.json", params=p, timeout=30).json()
    if d.get("status") != "000":
        return {}
    out: dict = {}
    for it in d.get("list", []) or []:
        aid = (it.get("account_id") or "").strip()
        nm = (it.get("account_nm") or "").strip()
        amt = _to_int(it.get("thstrm_amount"))
        if amt is None:
            continue
        target = ID_MAP.get(aid)
        if target is None and aid.startswith("-"):
            for frag, t in NM_FALLBACK:
                if frag in nm:
                    target = t
                    break
        if target and target not in out:   # 첫 등장 우선(중복 행 방지)
            out[target] = amt
    return out


def build_long_rows(key: str, corp: str, company: str, years: list[int]) -> pd.DataFrame:
    recs = []
    for year in years:
        for fs_div, gubun in (("OFS", "별도"), ("CFS", "연결")):
            vals = fetch_year(key, corp, year, fs_div)
            for (sj, acct), v in vals.items():
                recs.append({"회사명": company, "사업연도": year, "구분": gubun,
                             "재무제표": sj, "계정": acct, "값": v})
    return pd.DataFrame(recs)


def rebuild_pivot(long_df: pd.DataFrame, gubun: str, sj: str) -> pd.DataFrame:
    sub = long_df[(long_df["구분"] == gubun) & (long_df["재무제표"] == sj)]
    if sub.empty:
        return pd.DataFrame()
    piv = (sub.groupby(["계정", "사업연도"])["값"].last().unstack("사업연도"))
    piv = piv.reindex(sorted(piv.columns), axis=1).reset_index()
    return piv


def main(argv):
    ap = argparse.ArgumentParser()
    ap.add_argument("--corp", required=True)
    ap.add_argument("--company", required=True)
    ap.add_argument("--base", required=True, help="기존 시계열 Excel (PDF 기반)")
    ap.add_argument("--years", nargs="+", type=int, required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--key", default=os.environ.get("DART_API_KEY", ""))
    args = ap.parse_args(argv[1:])

    # .env 자동 로드 (extractor와 동일 동작)
    if not args.key:
        envp = HERE / ".env"
        if envp.exists():
            for line in envp.read_text(encoding="utf-8").splitlines():
                if line.strip().startswith("DART_API_KEY="):
                    args.key = line.split("=", 1)[1].strip().strip('"').strip("'")
    if not args.key:
        print("[!] DART_API_KEY 필요"); sys.exit(1)

    base = pd.ExcelFile(args.base)
    sheets = {s: pd.read_excel(args.base, sheet_name=s) for s in base.sheet_names}

    base_long = sheets["long_data"]
    api_long = build_long_rows(args.key, args.corp, args.company, args.years)
    print(f"API 수집: {len(api_long)}행 (연도 {args.years})")

    # 병합: 같은 (회사,연도,구분,재무제표,계정) 중복 시 API(뒤) 우선 last
    merged = pd.concat([base_long, api_long], ignore_index=True)
    merged = (merged.groupby(["회사명", "사업연도", "구분", "재무제표", "계정"], as_index=False)["값"]
                    .last())
    merged = merged.sort_values(["회사명", "구분", "재무제표", "사업연도", "계정"]).reset_index(drop=True)
    sheets["long_data"] = merged

    # 피벗 6종 재생성
    for gubun in ("별도", "연결"):
        for sj in ("재무상태표", "포괄손익계산서", "현금흐름표"):
            name = f"{gubun}_{sj}"
            piv = rebuild_pivot(merged, gubun, sj)
            if not piv.empty:
                sheets[name] = piv

    Path(args.out).parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(args.out, engine="openpyxl") as w:
        # long_data + 피벗 먼저, raw_* 는 기존 그대로 뒤에
        order = (["long_data"]
                 + [f"{g}_{s}" for g in ("별도", "연결") for s in ("재무상태표", "포괄손익계산서", "현금흐름표")]
                 + [s for s in sheets if s.startswith("raw_")])
        for s in order:
            if s in sheets:
                sheets[s].to_excel(w, sheet_name=s, index=False)
    print(f"WROTE: {args.out}")
    yrs = sorted(merged[merged['회사명'] == args.company]['사업연도'].unique())
    print(f"{args.company} 연도 범위: {yrs}")


if __name__ == "__main__":
    main(sys.argv)
