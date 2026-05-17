# -*- coding: utf-8 -*-
"""
DART OpenAPI + dart.fss.or.kr 페이지 스크래핑 래퍼.

기능:
- corpCode.xml 다운/캐시
- 회사명 키워드 검색
- 감사보고서/연결감사보고서 공시 목록 조회
- 공시 PDF 첨부파일 다운로드

DART OpenAPI 자체는 PDF 첨부 다운로드를 지원하지 않아, 공시 뷰어 페이지
(https://dart.fss.or.kr/dsaf001/main.do?rcpNo=...) 에서 dcmNo를 파싱한 뒤
PDF 다운 URL을 호출하는 방식을 쓴다.
"""
from __future__ import annotations

import io
import re
import sys
import time
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, Optional

import requests


# ---------- corpCode ----------

OPENDART = "https://opendart.fss.or.kr/api"


def ensure_corpcode(api_key: str, cache_xml: Path | str) -> Path:
    """corpCode.xml 다운/캐시. 이미 있으면 그대로 반환."""
    cache_xml = Path(cache_xml)
    if cache_xml.exists():
        return cache_xml
    cache_xml.parent.mkdir(parents=True, exist_ok=True)
    r = requests.get(f"{OPENDART}/corpCode.xml", params={"crtfc_key": api_key}, timeout=60)
    r.raise_for_status()
    if r.headers.get("content-type", "").startswith("application/json"):
        raise RuntimeError(f"corpCode API 에러 응답: {r.text[:500]}")
    with zipfile.ZipFile(io.BytesIO(r.content)) as zf:
        names = zf.namelist()
        # zip 안에 CORPCODE.xml 한 개만 들어 있음
        target = next((n for n in names if n.lower().endswith(".xml")), None)
        if not target:
            raise RuntimeError(f"zip에 XML이 없음: {names}")
        cache_xml.write_bytes(zf.read(target))
    return cache_xml


def find_companies(keyword: str, corpcode_xml: Path | str) -> list[dict]:
    """회사명에 keyword를 포함하는 후보들. 상장사 우선, 정확일치 우선 정렬."""
    keyword = keyword.strip()
    tree = ET.parse(str(corpcode_xml))
    cands = []
    for item in tree.getroot().iter("list"):
        name = (item.findtext("corp_name") or "").strip()
        if keyword in name:
            cands.append({
                "corp_code": (item.findtext("corp_code") or "").strip(),
                "corp_name": name,
                "stock_code": (item.findtext("stock_code") or "").strip(),
                "modify_date": (item.findtext("modify_date") or "").strip(),
            })
    # 정렬: 정확일치 → 상장사 → 짧은 이름
    def rank(c):
        exact = 0 if c["corp_name"] == keyword else 1
        listed = 0 if c["stock_code"].strip() else 1
        return (exact, listed, len(c["corp_name"]))
    cands.sort(key=rank)
    return cands


# ---------- 공시 목록 ----------

@dataclass
class Disclosure:
    rcept_no: str
    report_nm: str
    rcept_dt: str   # YYYYMMDD
    corp_name: str
    kind: str       # "별도" or "연결"

    def fiscal_hint_year(self) -> int:
        """공시 접수일 기준 사업연도 추정 (실제 사업연도는 PDF 파싱이 정확)."""
        return int(self.rcept_dt[:4]) - 1


def list_audit_disclosures(
    api_key: str,
    corp_code: str,
    bgn_de: str = "20100101",
    end_de: Optional[str] = None,
    sleep_sec: float = 0.15,
) -> list[Disclosure]:
    """감사보고서 + 연결감사보고서 공시 목록 (페이지네이션 처리).

    pblntf_ty=F는 "외부감사관련" 카테고리 전체. report_nm으로 감사보고서/연결감사보고서만 필터.
    """
    if end_de is None:
        from datetime import date
        end_de = date.today().strftime("%Y%m%d")

    out: list[Disclosure] = []
    page_no = 1
    while True:
        params = {
            "crtfc_key": api_key,
            "corp_code": corp_code,
            "bgn_de": bgn_de,
            "end_de": end_de,
            "pblntf_ty": "F",          # 외부감사관련
            "page_no": page_no,
            "page_count": 100,
        }
        r = requests.get(f"{OPENDART}/list.json", params=params, timeout=30)
        r.raise_for_status()
        data = r.json()
        status = data.get("status")
        if status == "013":   # 데이터 없음
            break
        if status != "000":
            raise RuntimeError(f"list.json 에러 status={status}: {data.get('message')}")

        for it in data.get("list", []):
            nm = (it.get("report_nm") or "").strip()
            # "감사보고서" 또는 "연결감사보고서"가 포함된 것만 (회계감사인 변경 등 제외)
            # 정정 공시도 포함(report_nm에 "[기재정정]감사보고서" 등으로 표시됨)
            if "감사보고서" not in nm:
                continue
            kind = "연결" if "연결" in nm else "별도"
            out.append(Disclosure(
                rcept_no=(it.get("rcept_no") or "").strip(),
                report_nm=nm,
                rcept_dt=(it.get("rcept_dt") or "").strip(),
                corp_name=(it.get("corp_name") or "").strip(),
                kind=kind,
            ))

        total_page = int(data.get("total_page") or 1)
        if page_no >= total_page:
            break
        page_no += 1
        time.sleep(sleep_sec)

    # 같은 사업연도에 정정공시가 있으면 가장 늦은 접수일자(=가장 최신)만 남기는 게 분석상 안전.
    # 다만 여기선 호출자에게 모두 노출하고, dedupe는 별도 헬퍼로 제공.
    return out


def dedupe_latest(disclosures: list[Disclosure]) -> list[Disclosure]:
    """같은 (사업연도 추정값, kind) 쌍에서 가장 최신 접수만 남김 (정정공시 우선)."""
    bucket: dict[tuple[int, str], Disclosure] = {}
    for d in sorted(disclosures, key=lambda x: x.rcept_dt):
        key = (d.fiscal_hint_year(), d.kind)
        bucket[key] = d   # 같은 키면 더 늦은 게 덮어씀
    return sorted(bucket.values(), key=lambda x: (x.kind, x.rcept_dt))


# ---------- PDF 다운로드 ----------

DART_VIEWER = "https://dart.fss.or.kr/dsaf001/main.do"
DART_PDF_INFO = "https://dart.fss.or.kr/pdf/download/main.do"      # 다운로드 안내 페이지
DART_PDF_DOWNLOAD = "https://dart.fss.or.kr/pdf/download/pdf.do"   # 실제 PDF 바이너리

# 뷰어 페이지 JS의 openPdfDownload(rcept_no, dcm_no) 호출에서 dcm_no를 뽑는다.
# 본문 XML용 dcmNo (node1['dcmNo']) 는 PDF와 다른 번호이므로 openPdfDownload 쪽이 정답.
_DCM_PATTERNS = [
    re.compile(r"openPdfDownload\(\s*['\"](\d+)['\"]\s*,\s*['\"](\d+)['\"]"),
    re.compile(r"viewDoc\(\s*['\"](\d+)['\"]\s*,\s*['\"](\d+)['\"]"),
]


def _extract_dcm_no(viewer_html: str, rcept_no: str) -> Optional[str]:
    """뷰어 페이지 HTML에서 PDF용 dcm_no 추출.

    openPdfDownload(rcpNo, dcmNo) 호출에서 매칭되는 쌍만 채택.
    (node1['dcmNo']는 본문 XML용 별도 번호라 사용하지 않음.)
    """
    for pat in _DCM_PATTERNS:
        for m in pat.finditer(viewer_html):
            if m.group(1) == rcept_no:
                return m.group(2)
            if m.group(2) == rcept_no:
                return m.group(1)
    return None


def download_pdf(
    rcept_no: str,
    dest_path: Path | str,
    *,
    session: Optional[requests.Session] = None,
) -> Path:
    """DART 공시 뷰어에서 PDF 첨부 다운로드.

    DART 다운로드는 세션 쿠키 + Referer 헤더가 필수다. 직접 pdf.do만 치면 빈 응답이 옴.
    실제 흐름:
      1) 뷰어 페이지(dsaf001/main.do) GET → HTML에서 dcm_no 추출 + 세션 쿠키 적재
      2) 다운로드 안내 페이지(pdf/download/main.do) GET → 추가 쿠키
      3) PDF 바이너리(pdf/download/pdf.do) GET + Referer 헤더

    이미 dest_path에 파일이 있으면 그대로 반환 (스킵).
    """
    dest_path = Path(dest_path)
    if dest_path.exists() and dest_path.stat().st_size > 1024:
        return dest_path
    dest_path.parent.mkdir(parents=True, exist_ok=True)

    sess = session or requests.Session()
    # requests의 기본 UA(python-requests/X) 는 DART가 빈 응답을 주므로 반드시 브라우저 UA로 덮어쓴다.
    sess.headers["User-Agent"] = (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0 Safari/537.36"
    )

    # 1) 뷰어 페이지 → dcm_no + 세션 쿠키
    r = sess.get(DART_VIEWER, params={"rcpNo": rcept_no}, timeout=30)
    r.raise_for_status()
    dcm_no = _extract_dcm_no(r.text, rcept_no)
    if not dcm_no:
        raise RuntimeError(
            f"dcm_no 추출 실패 (rcept_no={rcept_no}). 응답 일부: {r.text[:500]}"
        )

    # 2) 다운로드 안내 페이지 (쿠키 갱신)
    sess.get(DART_PDF_INFO, params={"rcp_no": rcept_no, "dcm_no": dcm_no}, timeout=30)

    # 3) PDF 바이너리
    r = sess.get(
        DART_PDF_DOWNLOAD,
        params={"rcp_no": rcept_no, "dcm_no": dcm_no},
        headers={"Referer": DART_PDF_INFO},
        timeout=120,
        stream=True,
    )
    r.raise_for_status()
    ctype = r.headers.get("content-type", "").lower()
    if "pdf" not in ctype and "octet-stream" not in ctype:
        snippet = r.content[:300]
        raise RuntimeError(
            f"PDF 응답 아님 (content-type={ctype}, rcept_no={rcept_no}, dcm_no={dcm_no}). "
            f"응답: {snippet!r}"
        )
    with open(dest_path, "wb") as f:
        for chunk in r.iter_content(chunk_size=64 * 1024):
            if chunk:
                f.write(chunk)
    if dest_path.stat().st_size < 1024:
        raise RuntimeError(f"PDF 크기가 너무 작음: {dest_path.stat().st_size}B")
    # PDF 시그니처 확인
    with open(dest_path, "rb") as f:
        head = f.read(4)
    if head != b"%PDF":
        raise RuntimeError(f"PDF 시그니처 불일치 (head={head!r}, rcept_no={rcept_no})")
    return dest_path


# ---------- CLI 디버그 진입점 ----------

def _print_help():
    print(__doc__)
    print("디버그 CLI:")
    print("  py dart_client.py search <키워드>          # 회사 검색")
    print("  py dart_client.py list <corp_code>          # 공시 목록")
    print("  py dart_client.py download <rcept_no> [out.pdf]   # PDF 다운로드 테스트")


def _main(argv: list[str]):
    import os
    if len(argv) < 2:
        _print_help()
        return
    api_key = os.environ.get("DART_API_KEY", "")
    cmd = argv[1]
    here = Path(__file__).resolve().parent
    corpcode_path = here / "data" / "CORPCODE.xml"

    if cmd == "search":
        if not api_key:
            print("[!] DART_API_KEY 환경변수 필요"); return
        ensure_corpcode(api_key, corpcode_path)
        for c in find_companies(argv[2], corpcode_path):
            print(c)
    elif cmd == "list":
        if not api_key:
            print("[!] DART_API_KEY 환경변수 필요"); return
        for d in list_audit_disclosures(api_key, argv[2]):
            print(d)
    elif cmd == "download":
        rcept_no = argv[2]
        out = Path(argv[3]) if len(argv) > 3 else here / "data" / f"{rcept_no}.pdf"
        p = download_pdf(rcept_no, out)
        print("saved:", p, "size:", p.stat().st_size)
    else:
        _print_help()


if __name__ == "__main__":
    _main(sys.argv)
