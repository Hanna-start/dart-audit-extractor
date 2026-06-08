# -*- coding: utf-8 -*-
"""1단계: 분기 재무 수집 (코스피, 본격).

기존 연간 데이터(financials)는 건드리지 않는다 — 별도 financials_q / collect_log_q.
분기보고서 4종(1분기·반기·3분기·연간)을 받는다. 손익/현금흐름은 '누적'값이라
받은 그대로 적재하고(헌법: 숫자 가공 없음), 순수 분기값 계산은 소비측(이 폴더 어댑터)에서.

기본 수집 계획:
  - 2022~2025: 4종 보고서 전부 (1분기 11013 / 반기 11012 / 3분기 11014 / 연간 11011)
  - 2026: 1분기(11013)만 (현재 그것만 공시됨)

체크포인트: (corp, year, reprt, fs_div) 단위로 collect_log_q에 기록 → 재실행 안전.
대상: data/kospi_index.json (build_kospi.py 산출).

사용:
  py screener/collect_quarterly.py                 # kospi_index 전체
  py screener/collect_quarterly.py --limit 100
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

from common import load_key, get_conn, DATA_DIR
import api_to_timeseries as api

KOSPI_INDEX = DATA_DIR / "kospi_index.json"
REPRT_NM = {"11013": "1분기", "11012": "반기", "11014": "3분기", "11011": "연간"}
FULL_REPRTS = ["11013", "11012", "11014", "11011"]

SCHEMA_Q = """
CREATE TABLE IF NOT EXISTS financials_q (
    corp_code  TEXT,
    bsns_year  INTEGER,
    reprt_code TEXT,     -- 11013=1Q / 11012=반기 / 11014=3Q / 11011=연간
    fs_div     TEXT,     -- OFS=별도 / CFS=연결
    재무제표    TEXT,
    계정        TEXT,
    값          INTEGER,  -- thstrm_amount: 손익/현금흐름=해당 분기 3개월, 재무상태=기말 시점값
    값_누적     INTEGER,  -- thstrm_add_amount: 손익/현금흐름 연초누적(분기보고서만; 연간·BS는 NULL)
    PRIMARY KEY (corp_code, bsns_year, reprt_code, fs_div, 재무제표, 계정)
);
CREATE TABLE IF NOT EXISTS collect_log_q (
    corp_code   TEXT,
    bsns_year   INTEGER,
    reprt_code  TEXT,
    fs_div      TEXT,
    status      TEXT,    -- ok / empty / error
    n_accounts  INTEGER,
    fetched_at  TEXT,
    PRIMARY KEY (corp_code, bsns_year, reprt_code, fs_div)
);
"""


def fetch_period(key: str, corp: str, year: int, reprt: str, fs_div: str) -> dict:
    """{(재무제표, 계정): (값, 값_누적)} — api_to_timeseries 매핑 재사용.

    값      = thstrm_amount (손익/현금흐름 분기보고서=해당 분기 3개월, BS=기말 시점)
    값_누적 = thstrm_add_amount (손익/현금흐름 연초누적; 연간·BS엔 없음 → None)
    """
    p = {"crtfc_key": key, "corp_code": corp, "bsns_year": str(year),
         "reprt_code": reprt, "fs_div": fs_div}
    d = requests.get(f"{api.OPEN}/fnlttSinglAcntAll.json", params=p, timeout=30).json()
    if d.get("status") != "000":
        return {}
    out: dict = {}
    for it in d.get("list", []) or []:
        aid = (it.get("account_id") or "").strip()
        nm = (it.get("account_nm") or "").strip()
        amt = api._to_int(it.get("thstrm_amount"))
        add = api._to_int(it.get("thstrm_add_amount"))
        if amt is None:
            continue
        target = api.ID_MAP.get(aid)
        if target is None and aid.startswith("-"):
            for frag, t in api.NM_FALLBACK:
                if frag in nm:
                    target = t
                    break
        if target and target not in out:
            out[target] = (amt, add)
    return out


def _already(conn, corp, year, reprt, fs_div, recheck_empty=False) -> bool:
    row = conn.execute(
        "SELECT status FROM collect_log_q WHERE corp_code=? AND bsns_year=? "
        "AND reprt_code=? AND fs_div=?", (corp, year, reprt, fs_div)).fetchone()
    if row is None:
        return False
    if recheck_empty:
        return row[0] == "ok"           # ok만 skip, empty는 재시도(지각 공시 대비)
    return row[0] in ("ok", "empty")


def filing_deadline(year: int, reprt: str) -> date:
    """한국 정기보고서 통상 제출 마감일."""
    return {
        "11013": date(year, 5, 15),       # 1분기
        "11012": date(year, 8, 14),       # 반기
        "11014": date(year, 11, 14),      # 3분기
        "11011": date(year + 1, 3, 31),   # 사업(연간)
    }[reprt]


def available_periods(start_year: int, today: date) -> list[tuple]:
    """오늘 기준 '제출 마감이 지나 공시됐을' 분기만 (미공시 미래분기 요청 안 함 → empty 오염 방지)."""
    periods = []
    for y in range(start_year, today.year + 1):
        for r in FULL_REPRTS:
            if filing_deadline(y, r) <= today:
                periods.append((y, r))
    return periods


def build_periods(years) -> list[tuple]:
    """수동 지정: 주어진 연도들의 4종 보고서 전부."""
    return [(y, r) for y in years for r in FULL_REPRTS]


def collect(corps, periods, key, conn, sleep=0.25, recheck_empty=False) -> dict:
    today = date.today().isoformat()
    stats = {"ok": 0, "empty": 0, "skip": 0, "error": 0}
    for c in corps:
        corp = c["corp_code"]
        for (year, reprt) in periods:
            for fs_div in ("CFS", "OFS"):
                if _already(conn, corp, year, reprt, fs_div, recheck_empty):
                    stats["skip"] += 1
                    continue
                try:
                    vals = fetch_period(key, corp, year, reprt, fs_div)
                except Exception as e:
                    with conn:
                        conn.execute("INSERT OR REPLACE INTO collect_log_q VALUES(?,?,?,?,?,?,?)",
                                     (corp, year, reprt, fs_div, "error", 0, today))
                    print(f"  [error] {corp} {year} {REPRT_NM[reprt]} {fs_div}: {e}")
                    stats["error"] += 1
                    time.sleep(sleep)
                    continue
                status = "ok" if vals else "empty"
                with conn:
                    for (sj, acct), (v, vadd) in vals.items():
                        conn.execute("INSERT OR REPLACE INTO financials_q VALUES(?,?,?,?,?,?,?,?)",
                                     (corp, year, reprt, fs_div, sj, acct, v, vadd))
                    conn.execute("INSERT OR REPLACE INTO collect_log_q VALUES(?,?,?,?,?,?,?)",
                                 (corp, year, reprt, fs_div, status, len(vals), today))
                stats[status] += 1
                time.sleep(sleep)
        print(f"  ✓ {c.get('corp_name', corp):24s} ({corp})")
    return stats


def main(argv):
    ap = argparse.ArgumentParser()
    ap.add_argument("--start-year", type=int, default=2022,
                    help="이 연도부터 (마감 지난 분기는 자동 포함)")
    ap.add_argument("--years", nargs="+", type=int, default=None,
                    help="수동 지정(미지정 시 자동: 오늘 기준 공시 마감 지난 모든 분기)")
    ap.add_argument("--recheck-empty", action="store_true",
                    help="empty로 기록된 것도 다시 시도(지각 공시 보완; 평소엔 불필요)")
    ap.add_argument("--limit", type=int, default=0, help="명단 앞에서 N개만(시총순)")
    ap.add_argument("--index", default=None,
                    help="대상 명단 json 경로(기본: kospi_index.json). 코스닥은 data/kosdaq_index.json")
    ap.add_argument("--key", default="")
    args = ap.parse_args(argv[1:])

    key = load_key(args.key)
    if not key:
        print("[!] DART_API_KEY 필요"); sys.exit(1)
    index_path = Path(args.index) if args.index else KOSPI_INDEX
    if not index_path.exists():
        print(f"[!] {index_path} 없음 — 먼저 build_*_universe.py 실행"); sys.exit(1)

    corps = json.loads(index_path.read_text(encoding="utf-8"))
    if args.limit:
        corps = corps[:args.limit]

    today = date.today()
    if args.years:
        periods = build_periods(args.years)            # 수동 override
    else:
        periods = available_periods(args.start_year, today)   # 자동 공시 달력 게이트

    conn = get_conn()
    with conn:
        conn.executescript(SCHEMA_Q)   # 별도 테이블만 추가 (기존 financials 불변)

    yr_span = sorted({y for y, _ in periods})
    print(f"오늘: {today} → 요청 대상은 '공시 마감 지난' 분기만 (미공시 미래분기 제외)")
    print(f"기간 {len(periods)}개 (연도 {yr_span[0]}~{yr_span[-1]}): " +
          ", ".join(f"{y}{REPRT_NM[r]}" for y, r in periods))
    print(f"분기 수집: {index_path.stem} {len(corps)}개사 × {len(periods)}기간 × 2(CFS/OFS)")
    print(f"예상 콜 ≈ {len(corps) * len(periods) * 2} (이미 받은 건 체크포인트로 skip)")
    stats = collect(corps, periods, key, conn, recheck_empty=args.recheck_empty)
    conn.close()
    print(f"\n수집 통계: {stats}")
    print("다음에 재실행하면 새로 공시된 분기만 추가로 받습니다(나머지 skip).")


if __name__ == "__main__":
    main(sys.argv)
