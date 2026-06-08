# -*- coding: utf-8 -*-
"""미국 분기 재무 수집 (SEC EDGAR companyfacts → financials_q).

한국 collect_quarterly.py의 미국판. 한국이 'DART fnlttSinglAcntAll(회사·분기·재무제표
단위 다중 콜)'이었다면, 미국은 'EDGAR companyfacts(회사당 단 1콜로 전 재무·전 기간)'다.
SEC는 API키·승인 불필요(User-Agent 이메일만), 일한도 없음, CDN 배포라 빠르고 가볍다.

핵심 변환(한국 누적차감 로직 그대로 재사용하려고 한국 스키마/계정명에 맞춰 적재):
  - US-GAAP 태그 → 한국 계정명(매출액·영업이익…)으로 번역 → financials_q.계정
  - 흐름 항목(매출·이익·현금흐름)은 한 회계연도의 YTD 누적이 전부 같은 start를 공유
    (예: 애플 FY25는 전부 2024-09-29 시작, 기간만 3·6·9·12개월) → DART 값_누적과 동일 개념.
    기간(개월) → 한국 reprt: 3=11013(Q1) 6=11012(반기) 9=11014(3Q) 12=11011(연간).
  - bsns_year = 회계연도(12개월 fact의 end 연도). 회계연도 말월이 12월이 아니어도(애플 9월,
    월마트 1월) 스크리너는 상대 분기로만 보므로 무관.
  - corp_code 자리 = CIK(10자리), fs_div = 'CFS'(미국은 연결 기준).

EDGAR 현실 대응:
  - 태그 전환: 회사가 옛 태그→새 태그로 바꿈(NVDA: Revenues가 주력, 옛 RevenueFromContract는 일부)
    → 후보 중 '최신·최다 커버리지' 태그 선택(첫 태그 아님).
  - Liabilities 태그 없는 회사(WMT·KO 등) → 부채 = 자산 − 자본(incl NCI)으로 유도.
  - 은행·보험(JPM·BRK 등)은 영업이익/매출원가 태그가 없어 품질단계에서 자연 탈락(한국 금융 제외와 동일).
  - 같은 기간이 fy=2025·2026 두 번 나오는 EDGAR 중복은 start/end 날짜 기준이라 무해(값 동일).

체크포인트: collect_log_q에 회사당 1행(bsns_year=0, reprt_code='USALL'). 재실행 시 ok는 skip,
  --refresh로 전 종목 재적재(새 분기 갱신; INSERT OR REPLACE라 안전).

사용:
  py screener/collect_us.py                    # us_index 전체(S&P500)
  py screener/collect_us.py --limit 20         # 앞 20개(표본 검증)
  py screener/collect_us.py --refresh          # 새 분기 갱신(전 종목 재적재)
"""
from __future__ import annotations

import argparse
import json
import sys
import time
from datetime import date
from pathlib import Path

import requests

if __name__ == "__main__":
    for _s in (sys.stdout, sys.stderr):
        try:
            if hasattr(_s, "reconfigure"):
                _s.reconfigure(encoding="utf-8")
        except Exception:
            pass

from common import get_conn, DATA_DIR

US_INDEX = DATA_DIR / "us_index.json"
UA = {"User-Agent": "Hanna Park jwnee2013@gmail.com"}
FACTS_URL = "https://data.sec.gov/api/xbrl/companyfacts/CIK{cik}.json"

# US-GAAP 태그 후보(폴백) → 한국 계정명. '최신·최다 커버리지' 태그를 고름.
FLOW_TAGS = {
    "매출액": ["RevenueFromContractWithCustomerExcludingAssessedTax", "Revenues",
             "RevenueFromContractWithCustomerIncludingAssessedTax", "SalesRevenueNet"],
    "영업이익": ["OperatingIncomeLoss"],
    "매출원가": ["CostOfGoodsAndServicesSold", "CostOfRevenue", "CostOfGoodsSold"],
    "영업활동현금흐름": ["NetCashProvidedByUsedInOperatingActivities",
                  "NetCashProvidedByUsedInOperatingActivitiesContinuingOperations"],
    "당기순이익": ["NetIncomeLoss", "ProfitLoss"],
}
EQUITY_TAGS = ["StockholdersEquityIncludingPortionAttributableToNoncontrollingInterest",
               "StockholdersEquity"]   # 총자본(지배+비지배) 우선
REPRT_BY_DUR = {3: "11013", 6: "11012", 9: "11014", 12: "11011"}
SJ = {"매출액": "손익계산서", "영업이익": "손익계산서", "매출원가": "손익계산서",
      "당기순이익": "손익계산서", "영업활동현금흐름": "현금흐름표",
      "부채총계": "재무상태표", "자본총계": "재무상태표", "자산총계": "재무상태표"}


def _months(start: str, end: str) -> int:
    y1, m1, d1 = map(int, start.split("-")); y2, m2, d2 = map(int, end.split("-"))
    mo = (y2 - y1) * 12 + (m2 - m1)
    if d2 - d1 > 20:
        mo += 1                      # 말일 경계(9/29~12/28 ≈ 3mo) 보정
    return mo


def _bucket(mo: int):
    for tgt in (3, 6, 9, 12):
        if tgt - 1 <= mo <= tgt + 1:
            return tgt
    return None


def _pick_tag(gaap: dict, candidates: list[str]):
    """후보 중 USD facts가 가장 최근까지·가장 많이 있는 태그의 facts(태그 전환 대응)."""
    best, best_key = None, None
    for t in candidates:
        node = gaap.get(t)
        if not node:
            continue
        facts = node.get("units", {}).get("USD")
        if not facts:
            continue
        max_end = max((f.get("end", "") for f in facts), default="")
        key = (max_end, len(facts))             # 최신 end 우선, 동률이면 fact 많은 것
        if best_key is None or key > best_key:
            best, best_key = facts, key
    return best


def _instant_map(gaap: dict, candidates: list[str]) -> dict:
    """시점(재무상태표) 항목: end date -> val. 나중 fact가 덮음(최신 정정 우선)."""
    facts = _pick_tag(gaap, candidates)
    out = {}
    if facts:
        for f in facts:
            en, v = f.get("end"), f.get("val")
            if en and v is not None:
                out[en] = v
    return out


def extract_rows(cik10: str) -> list[tuple]:
    """companyfacts → [(bsns_year, reprt_code, 재무제표, 계정, 값, 값_누적)]."""
    j = requests.get(FACTS_URL.format(cik=cik10), headers=UA, timeout=60).json()
    gaap = j.get("facts", {}).get("us-gaap", {})
    if not gaap:
        return []

    # 1) 흐름: 후보 중 주력 태그의 YTD를 (FYstart, dur)로. FYstart는 6·9·12mo의 start만 인정.
    raw = {}
    for acct, tags in FLOW_TAGS.items():
        facts = _pick_tag(gaap, tags)
        if not facts:
            continue
        lst = []
        for f in facts:
            st, en, v = f.get("start"), f.get("end"), f.get("val")
            if not st or not en or v is None:
                continue
            dur = _bucket(_months(st, en))
            if dur:
                lst.append((st, en, dur, v))
        raw[acct] = lst

    fy_starts = {st for lst in raw.values() for (st, en, dur, v) in lst if dur in (6, 9, 12)}
    fy_by_start, flow_ytd = {}, {}
    for acct, lst in raw.items():
        for (st, en, dur, v) in lst:
            if st not in fy_starts:        # 연중간 시작 3개월 fact 등 제외
                continue
            flow_ytd[(acct, st, dur)] = v
            d = fy_by_start.setdefault(st, {"year": None, "end": {}})
            d["end"][dur] = en
            if dur == 12:
                d["year"] = int(en[:4])
    for st, d in fy_by_start.items():
        if d["year"] is None:              # 연간 미발표(진행중) → end 연도로 추정
            d["year"] = int((max(d["end"].values()) if d["end"] else st)[:4])

    # 2) 시점: 자산·자본·부채(없으면 자산-자본 유도)
    assets_at = _instant_map(gaap, ["Assets"])
    equity_at = _instant_map(gaap, EQUITY_TAGS)
    liab_at = _instant_map(gaap, ["Liabilities"])

    def debt_at(end):
        if end in liab_at:
            return liab_at[end]
        if end in assets_at and end in equity_at:
            return assets_at[end] - equity_at[end]    # 부채=자산-자본
        return None

    # 3) (bsns_year, reprt_code) 레코드 조립
    rows = []
    for st, d in fy_by_start.items():
        year = d["year"]
        for dur, reprt in REPRT_BY_DUR.items():
            enddate = d["end"].get(dur)
            if enddate is None:
                continue
            for acct in FLOW_TAGS:
                v = flow_ytd.get((acct, st, dur))
                if v is None:
                    continue
                if reprt == "11011":
                    rows.append((year, reprt, SJ[acct], acct, v, None))      # 연간=값
                else:
                    rows.append((year, reprt, SJ[acct], acct, v, v))         # 분기=값_누적(값에도 동일)
            for acct, val in (("자산총계", assets_at.get(enddate)),
                              ("자본총계", equity_at.get(enddate)),
                              ("부채총계", debt_at(enddate))):
                if val is not None:
                    rows.append((year, reprt, SJ[acct], acct, val, None))
    return rows


def main(argv):
    ap = argparse.ArgumentParser()
    ap.add_argument("--index", default=str(US_INDEX), help="대상 명단 json(기본 us_index.json)")
    ap.add_argument("--limit", type=int, default=0, help="명단 앞에서 N개만(표본 검증)")
    ap.add_argument("--refresh", action="store_true", help="ok도 재적재(새 분기 갱신)")
    ap.add_argument("--sleep", type=float, default=0.12, help="콜 간 간격(초). SEC 초당10회 준수")
    args = ap.parse_args(argv[1:])

    index_path = Path(args.index)
    if not index_path.exists():
        print(f"[!] {index_path} 없음 — 먼저 build_us_universe.py 실행"); sys.exit(1)
    corps = json.loads(index_path.read_text(encoding="utf-8"))
    if args.limit:
        corps = corps[:args.limit]

    conn = get_conn()
    today = date.today().isoformat()
    stats = {"ok": 0, "empty": 0, "skip": 0, "error": 0}
    print(f"미국 분기 수집: {index_path.stem} {len(corps)}개사 (EDGAR companyfacts, 회사당 1콜)")

    for c in corps:
        cik = c["cik"]; name = c.get("name", cik); tk = c.get("ticker", "")
        done = conn.execute(
            "SELECT status FROM collect_log_q WHERE corp_code=? AND bsns_year=0 AND reprt_code='USALL'",
            (cik,)).fetchone()
        if done and done[0] == "ok" and not args.refresh:
            stats["skip"] += 1
            continue
        try:
            rows = extract_rows(cik)
        except Exception as e:
            with conn:
                conn.execute("INSERT OR REPLACE INTO collect_log_q VALUES(?,?,?,?,?,?,?)",
                             (cik, 0, "USALL", "CFS", "error", 0, today))
            print(f"  [error] {tk:6s} {name[:24]}: {e}")
            stats["error"] += 1
            time.sleep(args.sleep)
            continue
        status = "ok" if rows else "empty"
        with conn:
            for (year, reprt, sj, acct, v, va) in rows:
                conn.execute("INSERT OR REPLACE INTO financials_q VALUES(?,?,?,?,?,?,?,?)",
                             (cik, year, reprt, "CFS", sj, acct, v, va))
            conn.execute("INSERT OR REPLACE INTO collect_log_q VALUES(?,?,?,?,?,?,?)",
                         (cik, 0, "USALL", "CFS", status, len(rows), today))
        stats[status] += 1
        print(f"  ✓ {tk:6s} {name[:28]:28s} ({len(rows)}행)")
        time.sleep(args.sleep)

    conn.close()
    print(f"\n수집 통계: {stats}")


if __name__ == "__main__":
    main(sys.argv)
