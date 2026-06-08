# -*- coding: utf-8 -*-
"""미국 S&P 500 정식 명단 생성 (FinanceDataReader + SEC EDGAR 기준).

한국 build_kosdaq_universe.py의 미국판. 한국이 'FDR 명단 + DART corp_code'였다면,
미국은 'FDR S&P500 명단 + SEC CIK'다. CIK는 EDGAR에서 회사 재무를 찾는 키
(한국 corp_code 자리). 결과를 data/us_index.json에 저장한다.

- 명단: FinanceDataReader('S&P500') — 현재 S&P500 구성종목(Symbol/Name/Sector).
- 티커→CIK: SEC company_tickers.json(공식, 키 불필요, User-Agent만).
- 클래스주 표기차(SEC 'BRK-B' vs FDR 'BRK.B') → 점·하이픈 제거 정규화로 매칭.
- CIK는 EDGAR 규약대로 10자리 0채움 문자열로 저장.

사용: py screener/build_us_universe.py
"""
from __future__ import annotations

import json
import sys

import requests

if __name__ == "__main__":
    for _s in (sys.stdout, sys.stderr):
        try:
            if hasattr(_s, "reconfigure"):
                _s.reconfigure(encoding="utf-8")
        except Exception:
            pass

from common import DATA_DIR

US_INDEX = DATA_DIR / "us_index.json"
SEC_TICKERS = "https://www.sec.gov/files/company_tickers.json"
# SEC는 본인 식별용 User-Agent(이메일)만 요구 — 발급/승인 절차 없음.
UA = {"User-Agent": "Hanna Park jwnee2013@gmail.com"}


def _norm(t: str) -> str:
    """클래스주 표기 정규화: 점·하이픈·공백 제거 후 대문자 (BRK.B==BRK-B==BRKB)."""
    return t.upper().replace(".", "").replace("-", "").replace(" ", "")


def load_cik_map() -> dict:
    """SEC company_tickers.json → {정규화티커: (cik10, SEC원본티커, 공식명)}.
    SEC 티커는 클래스주를 'BRK-B'(하이픈)로 표기 = yfinance 조회 형식과 동일 → 그대로 저장."""
    j = requests.get(SEC_TICKERS, headers=UA, timeout=30).json()
    m = {}
    for row in j.values():
        tk = str(row.get("ticker", "")).strip()
        cik = str(row.get("cik_str", "")).strip()
        if not tk or not cik:
            continue
        m[_norm(tk)] = (cik.zfill(10), tk, row.get("title", ""))
    return m


def main():
    import FinanceDataReader as fdr

    df = fdr.StockListing("S&P500")
    print(f"FDR S&P500 구성종목: {len(df)}")

    cik_map = load_cik_map()
    print(f"SEC 티커→CIK 매핑: {len(cik_map)}개사 로드")

    rows, unmatched = [], []
    for _, r in df.iterrows():
        ticker = str(r["Symbol"]).strip()
        hit = cik_map.get(_norm(ticker))
        if not hit:
            unmatched.append(ticker)
            continue
        cik10, sec_ticker, sec_name = hit
        rows.append({
            "cik": cik10,                                   # EDGAR 조회 키(한국 corp_code 자리)
            "ticker": sec_ticker,                           # yfinance 조회 키(SEC 표기='BRK-B' 형식)
            "name": str(r.get("Name", "")) or sec_name,
            "sector": str(r.get("Sector", "")),
            "industry": str(r.get("Industry", "")),
        })

    US_INDEX.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"미국 S&P500 명단: {len(rows)}개사 → {US_INDEX.name}")
    if unmatched:
        print(f"CIK 미매칭 {len(unmatched)}개(SEC 명단에 티커 없음): {', '.join(unmatched)}")
    print("상위 5:", ", ".join(f"{x['name']}({x['ticker']}/CIK{x['cik']})" for x in rows[:5]))


if __name__ == "__main__":
    main()
