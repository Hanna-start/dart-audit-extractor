# -*- coding: utf-8 -*-
"""파일럿 한계 진단 — 4대 한계를 '실제 숫자'로 드러낸다.

수집된 sqlite를 읽어 회사별로:
  L1 금융업    : induty_code 64~66 → 일반 비율 왜곡 섹터
  L2 계정 결측 : 핵심 계정(매출/영업이익/순이익/자산·부채·자본총계) 누락
  L3 연결/별도 : CFS·OFS 가용성 분포 (어느 쪽으로 시계열을 잡을지)
  L4 데이터품질: BS 등식(자산=부채+자본) 위배 → 파싱·정합 깨진 회사

결정성 분리: 검증 로직은 축 1 validate.py의 등식을 그대로 차용(재계산 X 원칙).

사용:
  py screener/diagnose.py
"""
from __future__ import annotations

import sys

if __name__ == "__main__":
    for _s in (sys.stdout, sys.stderr):
        try:
            if hasattr(_s, "reconfigure"):
                _s.reconfigure(encoding="utf-8")
        except Exception:
            pass

from common import get_conn

TOL = 100   # validate.py와 동일 허용오차

# 핵심 계정 (없으면 스크리닝 불가)
CORE = {
    "재무상태표": ["자산총계", "부채총계", "자본총계"],
    "포괄손익계산서": ["매출액", "영업이익", "당기순이익"],
}


def _load(conn):
    companies = {r[0]: {"name": r[1], "cls": r[2], "induty": r[3], "fin": r[4]}
                 for r in conn.execute(
                     "SELECT corp_code, corp_name, corp_cls, induty_code, is_financial FROM companies")}
    fin = {}   # (corp, year, fs_div) → {(재무제표,계정): 값}
    for corp, yr, fs, sj, acct, val in conn.execute(
            "SELECT corp_code, bsns_year, fs_div, 재무제표, 계정, 값 FROM financials"):
        fin.setdefault((corp, yr, fs), {})[(sj, acct)] = val
    return companies, fin


def _best_fs(fin, corp, year):
    """연결 우선, 없으면 별도. (있는 쪽의 dict, 라벨) 반환."""
    cfs = fin.get((corp, year, "CFS"))
    if cfs:
        return cfs, "CFS"
    ofs = fin.get((corp, year, "OFS"))
    if ofs:
        return ofs, "OFS"
    return None, None


def diagnose():
    conn = get_conn()
    companies, fin = _load(conn)

    # 시도한 전체 corp (collect_log 기준) vs 실제 재무가 들어온 corp
    attempted = sorted({r[0] for r in conn.execute(
        "SELECT DISTINCT corp_code FROM collect_log")})
    collected = sorted({k[0] for k in fin.keys()})
    no_data = [c for c in attempted if c not in set(collected)]
    years = sorted({k[1] for k in fin.keys()})
    print(f"=== 파일럿 진단: 시도 {len(attempted)}사 / 재무 가용 {len(collected)}사, 연도 {years} ===\n")

    # L0: universe 노이즈 — 상장코드는 있으나 fnlttSinglAcntAll 무응답 (상폐·외국·신규)
    if no_data:
        print(f"⛔ L0 재무 무응답 {len(no_data)}사 (universe 노이즈):")
        for corp in no_data:
            info = companies.get(corp, {"name": corp, "cls": "?", "induty": "?"})
            print(f"     {info['name']:24s} cls={info.get('cls') or '?'} "
                  f"induty={info.get('induty') or '?'}")
        print()

    L1_financial, L2_missing, L4_bs_break = [], [], []
    avail = {"CFS만": 0, "OFS만": 0, "둘다": 0, "없음": 0}

    for corp in collected:
        info = companies.get(corp, {"name": corp, "fin": 0})
        name = info["name"]
        fin_flag = info.get("fin")

        # L3: 가용성
        has_cfs = any((corp, y, "CFS") in fin and fin[(corp, y, "CFS")] for y in years)
        has_ofs = any((corp, y, "OFS") in fin and fin[(corp, y, "OFS")] for y in years)
        key = ("둘다" if has_cfs and has_ofs else
               "CFS만" if has_cfs else "OFS만" if has_ofs else "없음")
        avail[key] += 1

        # L1: 금융업
        if fin_flag:
            L1_financial.append(name)

        # L2 + L4: 최신 가용 연도 기준
        miss_acc, bs_break_years = set(), []
        for y in years:
            d, label = _best_fs(fin, corp, y)
            if not d:
                continue
            present = {acct for (_sj, acct) in d.keys()}
            for sj, accts in CORE.items():
                for a in accts:
                    if a not in present:
                        miss_acc.add(a)
            a = d.get(("재무상태표", "자산총계"))
            l = d.get(("재무상태표", "부채총계"))
            e = d.get(("재무상태표", "자본총계"))
            if a is not None and l is not None and e is not None:
                if abs(a - l - e) >= TOL:
                    bs_break_years.append((y, label, a - l - e))

        flags = []
        if fin_flag:
            flags.append("금융")
        if miss_acc:
            flags.append("결측:" + ",".join(sorted(miss_acc)))
            L2_missing.append((name, sorted(miss_acc)))
        if bs_break_years:
            flags.append("BS깨짐:" + ",".join(str(y) for y, _, _ in bs_break_years))
            L4_bs_break.append((name, bs_break_years))

        mark = "⚠ " if flags else "  "
        print(f"{mark}{name:24s} [{key:5s}] {'; '.join(flags) if flags else 'OK'}")

    # ── 요약 ──
    print("\n" + "=" * 60)
    print("한계 정의 (파일럿 실측)")
    print("=" * 60)
    print(f"L0 재무 무응답(universe) : {len(no_data)}개사 (상폐 cls=E·외국·신규상장)")
    print(f"L1 금융업(비율 왜곡)     : {len(L1_financial)}개사  {L1_financial}")
    print(f"L2 핵심계정 결측         : {len(L2_missing)}개사")
    for nm, acc in L2_missing:
        print(f"     - {nm}: {acc}")
    print(f"L3 연결/별도 가용성 분포 : {avail}")
    print(f"L4 BS등식 위배(품질)     : {len(L4_bs_break)}개사")
    for nm, brk in L4_bs_break:
        print(f"     - {nm}: {[(y, fs, f'{d:,}') for y, fs, d in brk]}")
    print("\n→ 위 4개를 각각 어떻게 처리할지 정의한 뒤 전체 적재로 진행.")
    conn.close()


if __name__ == "__main__":
    diagnose()
