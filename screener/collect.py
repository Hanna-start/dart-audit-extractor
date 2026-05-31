# -*- coding: utf-8 -*-
"""② 배치 수집 — fnlttSinglAcntAll(재무) + company.json(업종) → sqlite.

장애 격리(체크포인트): (corp, year, fs_div) 단위로 collect_log에 기록 →
이미 받은 건 skip, 중단 시 그 지점부터 resume(Ctrl-C/네트워크 끊김 안전).
헌법: 숫자 가공 없음 — 받은 값을 그대로 적재.

주의: DART 일 한도 도달 시 fetch_year가 빈 dict를 줘 'empty'로 기록된다.
파일럿(~120콜)에선 무관하나, --all 전체 적재 전에 한도 status(020/021)를
직접 감지해 중단·재개하는 처리를 붙여야 한다. (TODO, 전체 적재 단계)

사용:
  py screener/collect.py --pilot 20 --years 2022 2023 2024
  py screener/collect.py --corps 00126380 005930 --years 2023 2024
  py screener/collect.py --all --years 2022 2023 2024     # 전체(주의: 수천 콜)
"""
from __future__ import annotations

import argparse
import sys
import time
from datetime import date

import requests

if __name__ == "__main__":
    for _s in (sys.stdout, sys.stderr):
        try:
            if hasattr(_s, "reconfigure"):
                _s.reconfigure(encoding="utf-8")
        except Exception:
            pass

from common import load_key, get_conn
import universe as uni

# fetch_year / ID_MAP 재사용 (축 1 자산)
import api_to_timeseries as api

OPEN = "https://opendart.fss.or.kr/api"


def fetch_company(key: str, corp: str) -> dict:
    """기업개황 — corp_cls(시장), induty_code(업종). 금융업 판정용."""
    p = {"crtfc_key": key, "corp_code": corp}
    try:
        d = requests.get(f"{OPEN}/company.json", params=p, timeout=30).json()
    except Exception:
        return {}
    if d.get("status") != "000":
        return {}
    return {
        "corp_cls": (d.get("corp_cls") or "").strip(),
        "induty_code": (d.get("induty_code") or "").strip(),
        "corp_name": (d.get("corp_name") or "").strip(),
    }


def _is_financial(induty_code: str) -> bool:
    """KSIC 대분류 K(금융·보험): 64~66. 일반 비율이 왜곡되는 섹터."""
    return induty_code[:2] in {"64", "65", "66"}


def _already(conn, corp, year, fs_div) -> bool:
    cur = conn.execute(
        "SELECT status FROM collect_log WHERE corp_code=? AND bsns_year=? AND fs_div=?",
        (corp, year, fs_div))
    row = cur.fetchone()
    return row is not None and row[0] in ("ok", "empty")


def collect(corps: list[dict], years: list[int], key: str, fs_divs=("CFS", "OFS"),
            sleep: float = 0.25) -> dict:
    conn = get_conn()
    today = date.today().isoformat()
    stats = {"company": 0, "ok": 0, "empty": 0, "skip": 0, "error": 0}

    for c in corps:
        corp = c["corp_code"]
        # 업종/시장 (있으면 skip 안 하고 갱신 — 가벼움)
        comp = fetch_company(key, corp)
        if comp:
            with conn:
                conn.execute(
                    "UPDATE companies SET corp_cls=?, induty_code=?, is_financial=?, "
                    "corp_name=COALESCE(NULLIF(?,''), corp_name), updated_at=? WHERE corp_code=?",
                    (comp["corp_cls"], comp["induty_code"],
                     1 if _is_financial(comp["induty_code"]) else 0,
                     comp["corp_name"], today, corp))
            stats["company"] += 1
        time.sleep(sleep)

        for year in years:
            for fs_div in fs_divs:
                if _already(conn, corp, year, fs_div):
                    stats["skip"] += 1
                    continue
                try:
                    vals = api.fetch_year(key, corp, year, fs_div)
                except Exception as e:
                    with conn:
                        conn.execute(
                            "INSERT OR REPLACE INTO collect_log VALUES(?,?,?,?,?,?)",
                            (corp, year, fs_div, "error", 0, today))
                    print(f"  [error] {corp} {year} {fs_div}: {e}")
                    stats["error"] += 1
                    time.sleep(sleep)
                    continue

                status = "ok" if vals else "empty"
                with conn:
                    for (sj, acct), v in vals.items():
                        conn.execute(
                            "INSERT OR REPLACE INTO financials VALUES(?,?,?,?,?,?)",
                            (corp, year, fs_div, sj, acct, v))
                    conn.execute("INSERT OR REPLACE INTO collect_log VALUES(?,?,?,?,?,?)",
                                 (corp, year, fs_div, status, len(vals), today))
                stats[status] += 1
                time.sleep(sleep)

        print(f"  ✓ {c['corp_name']:24s} ({corp})  "
              f"{'금융' if comp and _is_financial(comp['induty_code']) else '일반'}")

    conn.close()
    return stats


def main(argv):
    ap = argparse.ArgumentParser()
    g = ap.add_mutually_exclusive_group(required=True)
    g.add_argument("--pilot", type=int, help="모집단 등간격 표본 N개")
    g.add_argument("--corps", nargs="+", help="corp_code 직접 지정")
    g.add_argument("--all", action="store_true", help="상장사 전체 (수천 콜, 주의)")
    ap.add_argument("--years", nargs="+", type=int, required=True)
    ap.add_argument("--only", choices=["OFS", "CFS"], help="별도/연결 한쪽만")
    ap.add_argument("--key", default="")
    args = ap.parse_args(argv[1:])

    key = load_key(args.key)
    if not key:
        print("[!] DART_API_KEY 필요"); sys.exit(1)

    uni.build_index()   # 인덱스/companies 보장
    if args.pilot:
        corps = uni.pilot_sample(args.pilot)
    elif args.all:
        corps = uni.load_index()
    else:
        idx = {c["corp_code"]: c for c in uni.load_index()}
        corps = [idx.get(cc, {"corp_code": cc, "corp_name": cc, "stock_code": ""})
                 for cc in args.corps]

    fs_divs = (args.only,) if args.only else ("CFS", "OFS")
    print(f"수집 대상 {len(corps)}개사 × 연도 {args.years} × {fs_divs}")
    stats = collect(corps, args.years, key, fs_divs=fs_divs)
    print(f"\n수집 통계: {stats}")
    print("다음: py screener/diagnose.py 로 한계 진단")


if __name__ == "__main__":
    main(sys.argv)
