# -*- coding: utf-8 -*-
"""미국 '거래소 상장 전체' 명단 생성 (S&P500 밖으로 확장).

build_us_universe.py(S&P500 전용)의 확장판. 한국에서 코스피 다음 코스닥으로 넓힌 것과
같은 의미의 미국 확장이다.

- 명단: FinanceDataReader NASDAQ + NYSE + AMEX 보통주(거래소 상장사). 시총 컬럼은 없음.
- 티커→CIK: SEC company_tickers.json(공식, 키 불필요, User-Agent만). EDGAR 조회 키.
- S&P500(us_index.json) CIK는 이미 수집됐으므로 제외 → us_all_index.json엔 'S&P500 밖'만.
  (collect_us.py 체크포인트도 CIK로 skip하므로 안전장치는 이중.)

전략: EDGAR는 일한도가 없으나 7천 개 재무를 다 받기 전에, 먼저 이 명단으로 거래대금(유동성)
필터를 적용해 수집 대상을 줄인다(screen_us_liquidity.py). 그래서 여기선 CIK·티커·거래소만 담는다.

사용: py screener/build_us_all_universe.py
"""
from __future__ import annotations

import json
import sys

if __name__ == "__main__":
    for _s in (sys.stdout, sys.stderr):
        try:
            if hasattr(_s, "reconfigure"):
                _s.reconfigure(encoding="utf-8")
        except Exception:
            pass

from common import DATA_DIR
from build_us_universe import _norm, load_cik_map, US_INDEX

US_ALL_INDEX = DATA_DIR / "us_all_index.json"
EXCHANGES = ["NASDAQ", "NYSE", "AMEX"]


def main():
    import FinanceDataReader as fdr

    # 1) 거래소 상장명단(보통주) 수집 + 거래소 태깅, 정규화티커로 dedup
    listed = {}                       # norm_ticker -> (sec표기후보, name, exchange)
    for ex in EXCHANGES:
        df = fdr.StockListing(ex)
        for _, r in df.iterrows():
            sym = str(r.get("Symbol", "")).strip()
            if not sym:
                continue
            nt = _norm(sym)
            if nt and nt not in listed:
                listed[nt] = (sym, str(r.get("Name", "")), ex)
        print(f"FDR {ex}: {len(df)}종목 (누적 dedup {len(listed)})")

    # 2) 티커→CIK 매핑(SEC), S&P500 제외
    cik_map = load_cik_map()
    print(f"SEC 티커→CIK 매핑: {len(cik_map)}개사 로드")
    sp500_ciks = {c["cik"] for c in json.loads(US_INDEX.read_text(encoding="utf-8"))}
    print(f"S&P500 제외 대상: {len(sp500_ciks)}개사")

    rows, unmatched, sp_skip = [], 0, 0
    for nt, (sym, name, ex) in listed.items():
        hit = cik_map.get(nt)
        if not hit:
            unmatched += 1
            continue
        cik10, sec_ticker, sec_name = hit
        if cik10 in sp500_ciks:
            sp_skip += 1
            continue
        rows.append({
            "cik": cik10,                         # EDGAR 조회 키
            "ticker": sec_ticker,                 # yfinance/표시 키(SEC 표기)
            "name": name or sec_name,
            "exchange": ex,
        })

    # CIK 중복 제거(클래스주 등 한 회사 여러 티커 → 첫 티커만; 재무는 회사 단위라 1콜이면 충분)
    seen, uniq = set(), []
    for x in rows:
        if x["cik"] in seen:
            continue
        seen.add(x["cik"]); uniq.append(x)

    US_ALL_INDEX.write_text(json.dumps(uniq, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"\n미국 거래소 상장(S&P500 밖) 명단: {len(uniq)}개사(고유 CIK) → {US_ALL_INDEX.name}")
    print(f"  (CIK 미매칭 {unmatched} · S&P500 중복제외 {sp_skip} · 티커 {len(rows)}→CIK dedup {len(uniq)})")
    print("상위 5:", ", ".join(f"{x['name'][:18]}({x['ticker']})" for x in uniq[:5]))


if __name__ == "__main__":
    main()
