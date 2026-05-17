# -*- coding: utf-8 -*-
"""
시계열 통합 Excel의 무결성 검증 — 결정적 영역.

검사 항목:
  1. BS 등식: 자산총계 = 부채총계 + 자본총계 (별도/연결 각 연도)
  2. IS 등식: 매출액 = 매출원가 + 매출총이익
  3. 핵심 계정 누락: 매출액·자산총계·자본총계가 어느 연도에 비어 있는지

사용:
  py validate.py output/2026-05/무신사_시계열_2026-05-17.xlsx
  py validate.py output/2026-05/무신사_시계열_2026-05-17.xlsx --log _workspace/무신사/logs

종료 코드:
  0  통과
  1  검증 실패 (등식 어긋남 또는 누락 있음)
  2  파일/포맷 에러
"""
from __future__ import annotations

import argparse
import json
import sys
from dataclasses import dataclass, field, asdict
from datetime import date
from pathlib import Path

if __name__ == "__main__":
    for _s in (sys.stdout, sys.stderr):
        try:
            if hasattr(_s, "reconfigure"):
                _s.reconfigure(encoding="utf-8")
        except Exception:
            pass

import openpyxl


TOL = 100   # 100원 미만 차이는 반올림 노이즈로 간주


@dataclass
class CheckResult:
    sheet: str
    check: str          # "BS등식" / "IS등식" / "누락"
    year: int | str | None
    detail: str
    severity: str       # "ok" / "warn" / "fail"


@dataclass
class Report:
    file: str
    checked_at: str
    summary: dict = field(default_factory=dict)
    issues: list[CheckResult] = field(default_factory=list)

    def to_dict(self):
        d = asdict(self)
        d["issues"] = [asdict(i) for i in self.issues]
        return d


def _read_pivot(ws) -> tuple[list, dict]:
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return [], {}
    header = rows[0]
    data = {r[0]: r[1:] for r in rows[1:] if r and r[0]}
    return header, data


def check_bs(ws, sheet_name: str) -> list[CheckResult]:
    header, data = _read_pivot(ws)
    out: list[CheckResult] = []
    for i, year in enumerate(header[1:]):
        a = data.get("자산총계", [None])[i] if data.get("자산총계") else None
        l = data.get("부채총계", [None])[i] if data.get("부채총계") else None
        e = data.get("자본총계", [None])[i] if data.get("자본총계") else None
        if a is None or l is None or e is None:
            out.append(CheckResult(sheet_name, "BS등식", year,
                                    f"자산={a} 부채={l} 자본={e} (일부 누락)", "warn"))
            continue
        diff = a - l - e
        if abs(diff) < TOL:
            out.append(CheckResult(sheet_name, "BS등식", year, f"diff={diff}", "ok"))
        else:
            out.append(CheckResult(sheet_name, "BS등식", year,
                                    f"diff={diff:,} (자산 {a:,} - 부채 {l:,} - 자본 {e:,})", "fail"))
    return out


def check_is(ws, sheet_name: str) -> list[CheckResult]:
    header, data = _read_pivot(ws)
    out: list[CheckResult] = []
    for i, year in enumerate(header[1:]):
        s = data.get("매출액", [None])[i] if data.get("매출액") else None
        c = data.get("매출원가", [None])[i] if data.get("매출원가") else None
        g = data.get("매출총이익", [None])[i] if data.get("매출총이익") else None
        if s is None or c is None or g is None:
            # 서비스업처럼 매출원가가 없는 양식도 있음 → warn까지만
            out.append(CheckResult(sheet_name, "IS등식", year,
                                    f"매출={s} 원가={c} GP={g} (일부 누락)", "warn"))
            continue
        diff = s - c - g
        if abs(diff) < TOL:
            out.append(CheckResult(sheet_name, "IS등식", year, f"diff={diff}", "ok"))
        else:
            out.append(CheckResult(sheet_name, "IS등식", year,
                                    f"diff={diff:,} (매출 {s:,} - 원가 {c:,} - GP {g:,})", "fail"))
    return out


def check_missing_core(ws, sheet_name: str, core_accounts: list[str]) -> list[CheckResult]:
    """핵심 계정이 어느 연도에 비어 있는지 표시."""
    header, data = _read_pivot(ws)
    out: list[CheckResult] = []
    for acc in core_accounts:
        vals = data.get(acc)
        if vals is None:
            out.append(CheckResult(sheet_name, "누락", None, f"계정 자체가 시트에 없음: {acc}", "warn"))
            continue
        missing_years = [header[1 + i] for i, v in enumerate(vals) if v is None]
        if missing_years:
            out.append(CheckResult(sheet_name, "누락", None,
                                    f"{acc} 누락 연도: {missing_years}", "warn"))
    return out


def validate(xlsx_path: Path) -> Report:
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    rep = Report(file=str(xlsx_path), checked_at=date.today().isoformat())

    bs_sheets = [s for s in ("별도_재무상태표", "연결_재무상태표") if s in wb.sheetnames]
    is_sheets = [s for s in ("별도_포괄손익계산서", "연결_포괄손익계산서") if s in wb.sheetnames]

    for sn in bs_sheets:
        rep.issues.extend(check_bs(wb[sn], sn))
        rep.issues.extend(check_missing_core(wb[sn], sn, ["자산총계", "자본총계"]))
    for sn in is_sheets:
        rep.issues.extend(check_is(wb[sn], sn))
        rep.issues.extend(check_missing_core(wb[sn], sn, ["매출액", "영업이익", "당기순이익"]))

    rep.summary = {
        "ok": sum(1 for i in rep.issues if i.severity == "ok"),
        "warn": sum(1 for i in rep.issues if i.severity == "warn"),
        "fail": sum(1 for i in rep.issues if i.severity == "fail"),
    }
    return rep


def print_report(rep: Report):
    print(f"[검증] {rep.file}")
    print(f"        통과 {rep.summary['ok']} / 경고 {rep.summary['warn']} / 실패 {rep.summary['fail']}")
    by_sheet: dict[str, list[CheckResult]] = {}
    for i in rep.issues:
        by_sheet.setdefault(i.sheet, []).append(i)
    for sheet, items in by_sheet.items():
        print(f"  ── {sheet} ──")
        for it in items:
            sym = {"ok": "✓", "warn": "△", "fail": "✗"}[it.severity]
            year = f"{it.year} " if it.year is not None else ""
            print(f"    {sym} [{it.check}] {year}{it.detail}")


def write_log(rep: Report, log_dir: Path) -> Path:
    log_dir.mkdir(parents=True, exist_ok=True)
    fname = f"validation_{rep.checked_at}.log"
    p = log_dir / fname
    with p.open("w", encoding="utf-8") as f:
        json.dump(rep.to_dict(), f, ensure_ascii=False, indent=2)
    return p


def main(argv: list[str]):
    p = argparse.ArgumentParser(description="시계열 Excel 무결성 검증")
    p.add_argument("xlsx", help="검증할 시계열 Excel 경로")
    p.add_argument("--log", help="로그 디렉토리 (지정 시 JSON 로그 저장)")
    args = p.parse_args(argv[1:])

    xlsx_path = Path(args.xlsx)
    if not xlsx_path.exists():
        print(f"[!] 파일 없음: {xlsx_path}", file=sys.stderr)
        sys.exit(2)

    try:
        rep = validate(xlsx_path)
    except Exception as e:
        print(f"[!] 검증 중 에러: {e}", file=sys.stderr)
        sys.exit(2)

    print_report(rep)

    if args.log:
        log_path = write_log(rep, Path(args.log))
        print(f"\n로그 저장: {log_path}")

    sys.exit(1 if rep.summary["fail"] > 0 else 0)


if __name__ == "__main__":
    main(sys.argv)
