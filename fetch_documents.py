# -*- coding: utf-8 -*-
"""DART 공시 원문(document.xml) → 섹션별 마크다운 저장.

사업보고서처럼 서술형 본문(사업의 개요·주요 제품·위험관리 등)을 텍스트로 보존한다.
fnlttSinglAcntAll은 정형 숫자만 주므로, 서술 본문은 이 엔드포인트로 받아야 한다.

사용:
  py fetch_documents.py --corp 01137727 --company 무신사 --years 2024 2025 \
     --report 사업보고서 --out _workspace/무신사/documents
"""
from __future__ import annotations

import argparse
import html
import io
import os
import re
import sys
import zipfile
from datetime import date
from pathlib import Path

import requests

HERE = Path(__file__).resolve().parent
OPEN = "https://opendart.fss.or.kr/api"


def _load_key(cli_key: str) -> str:
    if cli_key:
        return cli_key
    k = os.environ.get("DART_API_KEY", "")
    if k:
        return k
    envp = HERE / ".env"
    if envp.exists():
        for line in envp.read_text(encoding="utf-8").splitlines():
            if line.strip().startswith("DART_API_KEY="):
                return line.split("=", 1)[1].strip().strip('"').strip("'")
    return ""


def find_reports(key, corp, report_kw, years):
    """report_nm 에 report_kw 와 (YYYY.12) 포함된 공시 찾기 → {year: (rcept_no, report_nm, rcept_dt)}."""
    p = {"crtfc_key": key, "corp_code": corp, "bgn_de": f"{min(years)}0101",
         "end_de": f"{max(years)+2}1231", "page_count": 100}
    lst = requests.get(f"{OPEN}/list.json", params=p, timeout=30).json().get("list", []) or []
    found = {}
    for it in lst:
        nm = (it.get("report_nm") or "")
        if report_kw not in nm:
            continue
        for y in years:
            if f"({y}.12)" in nm and y not in found:
                found[y] = (it["rcept_no"].strip(), nm.strip(), it.get("rcept_dt", ""))
    return found


def _clean(s: str) -> str:
    s = re.sub(r"<[^>]+>", "", s)
    s = s.replace("&cr;", " ").replace("&nbsp;", " ").replace("&#160;", " ")
    s = html.unescape(s)
    return re.sub(r"[ \t]+", " ", s).strip()


def _table_to_md(tbl: str) -> str:
    rows = re.findall(r"<TR\b[^>]*>(.*?)</TR>", tbl, re.S | re.I)
    out = []
    for r in rows:
        cells = re.findall(r"<T[HDE]\b[^>]*>(.*?)</T[HDE]>", r, re.S | re.I)
        cells = [_clean(c).replace("\n", " ") or " " for c in cells]
        if cells:
            out.append("| " + " | ".join(cells) + " |")
    if not out:
        return ""
    # 헤더 구분선 (첫 행 기준)
    ncol = out[0].count("|") - 1
    out.insert(1, "| " + " | ".join(["---"] * ncol) + " |")
    return "\n".join(out)


def _heading_level(title: str) -> int:
    t = title.strip()
    if t.startswith("【"):
        return 2
    if re.match(r"^[IVXLC]+\.", t):
        return 2
    if re.match(r"^\d+-\d+", t):
        return 4
    if re.match(r"^\d+\.", t):
        return 3
    return 3


def xml_to_markdown(raw: str) -> str:
    # 1) 테이블 먼저 마크다운으로 치환 (자리 보존)
    def tbl_repl(m):
        md = _table_to_md(m.group(0))
        return f"\n\n{md}\n\n" if md else "\n"
    raw = re.sub(r"<TABLE\b[^>]*>.*?</TABLE>", tbl_repl, raw, flags=re.S | re.I)
    # 2) 제목
    def title_repl(m):
        txt = _clean(m.group(1))
        if not txt:
            return ""
        return f"\n\n{'#' * _heading_level(txt)} {txt}\n\n"
    raw = re.sub(r"<TITLE\b[^>]*>(.*?)</TITLE>", title_repl, raw, flags=re.S | re.I)
    # 3) 문단
    raw = re.sub(r"<P\b[^>]*>(.*?)</P>", lambda m: _clean(m.group(1)) + "\n\n", raw, flags=re.S | re.I)
    # 4) 페이지 구분/남은 태그 제거
    raw = re.sub(r"<[^>]+>", " ", raw)
    raw = raw.replace("&cr;", " ").replace("&nbsp;", " ").replace("&#160;", " ")
    raw = html.unescape(raw)
    # 5) 공백 정리: 표 라인은 보존, 일반 줄만 정리
    lines = []
    for ln in raw.splitlines():
        ln = ln.rstrip()
        if ln.startswith("|"):
            lines.append(ln)
        else:
            ln = re.sub(r"[ \t]+", " ", ln).strip()
            lines.append(ln)
    text = "\n".join(lines)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def main(argv):
    ap = argparse.ArgumentParser()
    ap.add_argument("--corp", required=True)
    ap.add_argument("--company", required=True)
    ap.add_argument("--years", nargs="+", type=int, required=True)
    ap.add_argument("--report", default="사업보고서")
    ap.add_argument("--out", required=True)
    ap.add_argument("--key", default="")
    args = ap.parse_args(argv[1:])

    key = _load_key(args.key)
    if not key:
        print("[!] DART_API_KEY 필요"); sys.exit(1)

    reports = find_reports(key, args.corp, args.report, args.years)
    if not reports:
        print(f"[!] '{args.report}' 공시를 찾지 못함 (연도 {args.years})"); sys.exit(1)

    out_dir = Path(args.out)
    out_dir.mkdir(parents=True, exist_ok=True)

    for year in sorted(reports):
        rcept_no, nm, rcept_dt = reports[year]
        r = requests.get(f"{OPEN}/document.xml",
                         params={"crtfc_key": key, "rcept_no": rcept_no}, timeout=120)
        z = zipfile.ZipFile(io.BytesIO(r.content))
        raw = z.read(z.namelist()[0]).decode("utf-8", "ignore")
        md_body = xml_to_markdown(raw)
        header = (f"# {args.company} {nm}\n\n"
                  f"- 접수일: {rcept_dt}  ·  rcept_no: {rcept_no}  ·  사업연도: {year}\n"
                  f"- 출처: DART OpenAPI document.xml  ·  추출일: {date.today().isoformat()}\n"
                  f"- 원문 가공: tag→markdown (서술 보존, 표는 텍스트 변환)\n\n---\n\n")
        dest = out_dir / f"{args.report}_{year}.md"
        dest.write_text(header + md_body, encoding="utf-8")
        print(f"WROTE: {dest}  ({len(header + md_body)//1024} KB, 본문 {len(md_body):,}자)")


if __name__ == "__main__":
    main(sys.argv)
