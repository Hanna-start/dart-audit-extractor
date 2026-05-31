# -*- coding: utf-8 -*-
"""
감사보고서 PDF → 재무제표 4종 Excel 변환 (범용)

DART에 공시되는 한국 회사 감사보고서/연결감사보고서의 본문 재무제표를 추출.
대상 양식: K-IFRS 또는 일반기업회계기준 양식의 표준 DART 감사보고서.

사용 (CLI):
  단일 PDF:   py pdf_financials.py "감사보고서.pdf"
  폴더 일괄:  py pdf_financials.py "폴더경로"

출력:
  단일: 입력 PDF와 같은 폴더에 [원본이름].xlsx (시트: 재무상태표/포괄손익/자본변동/현금흐름 + meta)
  폴더: 폴더/추출결과/ 아래에 PDF별 xlsx + 통합 시계열 xlsx

라이브러리로 사용:
  from pdf_financials import extract_financials, save_excel
  fs = extract_financials("감사보고서.pdf")   # dict[str, DataFrame]
  save_excel(fs, "out.xlsx")
"""
from __future__ import annotations

import io
import re
import sys
import json
from pathlib import Path
from typing import Optional

import pandas as pd
import pdfplumber


# 콘솔 한글 깨짐 방지: 단독 실행시에만 reconfigure. 라이브러리로 import될 때는 호출자가 처리.
if __name__ == "__main__":
    for _s in (sys.stdout, sys.stderr):
        try:
            if hasattr(_s, "reconfigure"):
                _s.reconfigure(encoding="utf-8")
        except Exception:
            pass


# 재무제표 4종(+α) 제목 — 페이지 첫 줄에 큰 글씨로 등장하면 새 섹션 시작
# 키: 표준 명칭, 값: 매칭에 쓰는 정규식 (공백 무시)
SECTION_TITLES = {
    "재무상태표":     re.compile(r"^(연\s*결\s*)?재\s*무\s*상\s*태\s*표\s*$"),
    "포괄손익계산서": re.compile(r"^(연\s*결\s*)?(포\s*괄\s*)?손\s*익\s*계\s*산\s*서\s*$"),
    "자본변동표":     re.compile(r"^(연\s*결\s*)?자\s*본\s*변\s*동\s*표\s*$"),
    "현금흐름표":     re.compile(r"^(연\s*결\s*)?현\s*금\s*흐\s*름\s*표\s*$"),
}
# 본문 재무제표 종료 신호 (이게 페이지 첫 줄에 나오면 종료)
END_TITLES = re.compile(r"^(주\s*석|재무제표에\s*대한\s*주석|연결재무제표에\s*대한\s*주석)")

# 단위 표기 추출
UNIT_RE = re.compile(r"\(단위\s*[:：]\s*([^\)]+)\)")

# 회사명/사업연도 추출용 패턴
# "주식회사 XX" 또는 "(주)XX" 만 잡고, "....주식회사" 같은 목차 노이즈는 제외
COMPANY_RE = re.compile(
    # "주식회사 XX와 그 종속기업" → "XX"만 잡도록 lookahead 추가
    r"주식회사\s+([가-힣A-Za-z0-9][가-힣A-Za-z0-9 ]{0,20}?)(?=\s*와\s+그\s+종속|[\s,\.\n]|$)"
    r"|\(주\)\s*([가-힣A-Za-z0-9]+)"
)
PERIOD_RE = re.compile(r"제\s*\d+\s*\(?[당전]?\)?\s*기")
YEAR_RE = re.compile(r"(\d{4})\s*년")
# 감사보고서 표지의 "YYYY년 M월 D일 까지" 패턴 — 보고기간 종료일 = 사업연도의 정의
END_DATE_RE = re.compile(r"(\d{4})\s*년\s*\d{1,2}\s*월\s*\d{1,2}\s*일\s*까지")


# ---------- 핵심 추출 로직 ----------

def _page_title(text: str) -> Optional[str]:
    """페이지 텍스트의 첫 줄이 재무제표 섹션 제목인지 판별 → 섹션 이름 반환"""
    if not text:
        return None
    # 페이지에서 의미 있는 첫 줄 (공백/짧은 줄 스킵)
    for raw in text.splitlines()[:3]:
        line = raw.strip()
        if not line:
            continue
        if END_TITLES.match(line):
            return "__END__"
        for name, pat in SECTION_TITLES.items():
            if pat.match(line):
                return name
        # 첫 의미 라인이 섹션 제목이 아니면 그 페이지는 섹션 시작 아님
        return None
    return None


def _normalize_number(s: str) -> Optional[float]:
    """'1,234,567', '(1,234)' 같은 문자열을 숫자로"""
    if s is None:
        return None
    t = str(s).strip()
    if t in ("", "-", "―", "—"):
        return 0.0
    # 음수: 괄호 또는 △/▲ 표기
    neg = False
    if (t.startswith("(") and t.endswith(")")) or t.startswith("△") or t.startswith("▲"):
        neg = True
        t = t.strip("()△▲ ")
    t = t.replace(",", "").replace(" ", "")
    try:
        v = float(t)
        return -v if neg else v
    except ValueError:
        return None


def _clean_table(raw_table: list[list]) -> pd.DataFrame:
    """pdfplumber 표 → 정리된 DataFrame
    - None을 빈 문자열로
    - 빈 행 제거
    - 첫 컬럼은 계정명, 나머지는 그대로 (필요시 호출자가 후처리)
    """
    if not raw_table:
        return pd.DataFrame()
    rows = []
    for r in raw_table:
        r = [("" if c is None else str(c).strip()) for c in r]
        if any(r):
            rows.append(r)
    if not rows:
        return pd.DataFrame()
    # 컬럼 수 통일
    max_cols = max(len(r) for r in rows)
    rows = [r + [""] * (max_cols - len(r)) for r in rows]
    df = pd.DataFrame(rows)
    return df


def _dedup_columns(cols: list) -> list:
    """중복/빈 컬럼명을 유일하게 만들기"""
    seen = {}
    out = []
    for i, c in enumerate(cols):
        name = (str(c).strip() or f"col_{i}")
        if name in seen:
            seen[name] += 1
            out.append(f"{name}__{seen[name]}")
        else:
            seen[name] = 0
            out.append(name)
    return out


def _attach_numeric_columns(df: pd.DataFrame) -> pd.DataFrame:
    """숫자처럼 보이는 컬럼을 _num 접미사로 추가 (원본 보존)"""
    if df.empty:
        return df
    out = df.copy()
    for i, col in enumerate(df.columns):
        if i == 0:
            continue  # 첫 컬럼(계정명) 스킵
        series = df.iloc[:, i]
        nums = series.map(_normalize_number)
        if nums.notna().sum() >= max(2, len(df) // 3):
            out[f"{col}_num"] = nums
    return out


def extract_meta(pdf_pages_text: list[str], pdf_name: str = "") -> dict:
    """앞쪽 페이지에서 회사명/사업연도/단위/구분(별도/연결) 추출"""
    meta = {"회사명": None, "사업연도": None, "전기연도": None, "구분": "별도"}
    head = "\n".join(pdf_pages_text[:10])

    # 회사명: 목차 라인은 제외하고 본문에서 검색
    body_lines = [
        ln.strip() for ln in head.splitlines()
        if ln.strip() and "..." not in ln  # 목차 라인 배제
    ]
    body = "\n".join(body_lines)
    m = COMPANY_RE.search(body)
    if m:
        # group(1)은 "주식회사 XX" 매칭, group(2)는 "(주)XX" 매칭
        name = m.group(1) or m.group(2)
        if name:
            meta["회사명"] = name.strip()

    # 사업연도: 표지의 "YYYY년 M월 D일 까지" (보고기간 종료일) 우선
    # 폴백: 모든 "YYYY년" 중 max (공시일·감사일이 결산일보다 큰 함정 있음)
    end_years = [int(y) for y in END_DATE_RE.findall(head)]
    if end_years:
        ys = sorted(set(end_years), reverse=True)
        meta["사업연도"] = ys[0]
        meta["전기연도"] = ys[1] if len(ys) >= 2 else None
    else:
        years = YEAR_RE.findall(head)
        if years:
            ys = sorted({int(y) for y in years})
            meta["사업연도"] = ys[-1] if ys else None
            meta["전기연도"] = ys[-2] if len(ys) >= 2 else None

    # 별도/연결 구분: 파일명 또는 본문 키워드로 판단
    name_lower = str(pdf_name)
    if "연결" in name_lower or head.count("연결재무") >= 3:
        meta["구분"] = "연결"

    return meta


def extract_financials(pdf_path: str | Path, verbose: bool = False) -> dict:
    """PDF에서 재무제표 4종을 추출.

    Returns:
        {
            "meta": {...},
            "재무상태표": DataFrame | None,
            "포괄손익계산서": DataFrame | None,
            "자본변동표": DataFrame | None,
            "현금흐름표": DataFrame | None,
            "_pages": {"재무상태표": [7,8], ...}  # 디버깅용
        }
    """
    pdf_path = Path(pdf_path)
    result = {
        "meta": {},
        "재무상태표": None,
        "포괄손익계산서": None,
        "자본변동표": None,
        "현금흐름표": None,
        "_pages": {},
        "_source": str(pdf_path),
    }

    with pdfplumber.open(pdf_path) as pdf:
        # 1) 페이지별 텍스트/표 캐싱
        pages_text = []
        pages_tables = []
        for page in pdf.pages:
            pages_text.append(page.extract_text() or "")
            try:
                pages_tables.append(page.extract_tables() or [])
            except Exception:
                pages_tables.append([])

        # 2) 메타 정보
        result["meta"] = extract_meta(pages_text, pdf_name=pdf_path.name)

        # 2-1) 스캔본/파싱실패 조기 감지 — 예외로 터뜨리지 않고 경고만 기록.
        # 재무제표 PDF는 표가 반드시 있어야 하므로 '표 0개'는 거의 확정적 신호.
        # 표지페이지(텍스트 적음) 오탐을 피하려고 문서 전체 신호를 복합 판정한다.
        n_pages = len(pages_text)
        total_chars = sum(len(t.strip()) for t in pages_text)
        doc_tables = sum(len(tb) for tb in pages_tables)
        avg_chars = total_chars / max(1, n_pages)
        warnings: list[str] = []
        if total_chars < 100 or (doc_tables == 0 and avg_chars < 50):
            warnings.append("이미지 스캔본 PDF로 추정됨 (텍스트·표 거의 미검출, OCR 필요)")
        elif doc_tables == 0:
            warnings.append("표가 전혀 감지되지 않음 (비표준 양식 또는 스캔본 가능성, 확인 필요)")
        result["_warnings"] = warnings
        result["meta"]["추출경고"] = "; ".join(warnings) if warnings else None
        if warnings and verbose:
            for w in warnings:
                print(f"  [WARNING] {pdf_path.name}: {w}")

        # 3) 섹션 경계 탐지
        # 페이지마다 첫 줄 제목 확인 → 섹션 시작 페이지 마킹
        section_starts = []  # [(page_idx, section_name)]
        for i, t in enumerate(pages_text):
            title = _page_title(t)
            if title:
                section_starts.append((i, title))

        if verbose:
            print(f"[{pdf_path.name}] 섹션 마킹:", section_starts)

        # 4) 각 섹션의 페이지 범위 결정
        # 같은 섹션이 연속해서 표시되거나, 다른 섹션/END가 나오면 종료
        ranges = {}  # section_name -> [page_idx, ...]
        for idx, (start_page, name) in enumerate(section_starts):
            if name == "__END__":
                continue
            # 다음 마커까지의 페이지가 이 섹션 범위
            next_marker = section_starts[idx + 1][0] if idx + 1 < len(section_starts) else len(pages_text)
            ranges.setdefault(name, []).extend(range(start_page, next_marker))

        result["_pages"] = {k: [p + 1 for p in v] for k, v in ranges.items()}

        # 5) 각 섹션 페이지의 표들을 합쳐서 DataFrame 생성
        for name, page_idxs in ranges.items():
            collected_rows = []
            header = None
            for p in page_idxs:
                tables = pages_tables[p]
                # 같은 페이지에 표가 여러 개면 가장 큰 것 우선 (계정 많은 본 표)
                tables_sorted = sorted(tables, key=lambda t: len(t), reverse=True)
                for t in tables_sorted[:1]:  # 페이지당 최상위 표 하나만
                    df = _clean_table(t)
                    if df.empty:
                        continue
                    if header is None:
                        # 첫 페이지의 첫 1~2행을 헤더로 사용 (보통 "과 목 | 주석 | 제 N(당) 기 | 제 N-1(전) 기")
                        header = df.iloc[0].tolist()
                        rows = df.iloc[1:]
                    else:
                        # 연장 페이지: 헤더 행이 또 있을 수 있음. 첫 행이 header와 비슷하면 스킵.
                        if df.iloc[0].tolist() == header:
                            rows = df.iloc[1:]
                        else:
                            rows = df
                    collected_rows.extend(rows.values.tolist())

            if not collected_rows:
                continue

            # 컬럼 수 통일
            max_cols = max(len(r) for r in collected_rows)
            if header is None or len(header) < max_cols:
                header = (header or []) + [f"col_{i}" for i in range(len(header or []), max_cols)]
            collected_rows = [list(r) + [""] * (max_cols - len(r)) for r in collected_rows]
            df = pd.DataFrame(collected_rows, columns=_dedup_columns(header[:max_cols]))

            # 단위 추출 (해당 섹션 첫 페이지 텍스트에서)
            unit_m = UNIT_RE.search(pages_text[page_idxs[0]])
            unit = unit_m.group(1).strip() if unit_m else None

            df = _attach_numeric_columns(df)
            df.attrs["unit"] = unit
            df.attrs["pages"] = [p + 1 for p in page_idxs]
            result[name] = df

    return result


# ---------- 출력 ----------

def save_excel(fs: dict, out_path: str | Path) -> Path:
    """추출 결과를 Excel로 저장 (시트별 재무제표)"""
    out_path = Path(out_path)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_path, engine="openpyxl") as xw:
        # meta 시트
        meta_rows = list(fs["meta"].items()) + [
            ("원본PDF", fs.get("_source", "")),
        ]
        for name in ["재무상태표", "포괄손익계산서", "자본변동표", "현금흐름표"]:
            df = fs.get(name)
            if df is not None:
                meta_rows.append((f"{name}_단위", df.attrs.get("unit")))
                meta_rows.append((f"{name}_원본페이지", str(df.attrs.get("pages"))))
        meta_df = pd.DataFrame(meta_rows, columns=["항목", "값"])
        meta_df.to_excel(xw, sheet_name="meta", index=False)

        for name in ["재무상태표", "포괄손익계산서", "자본변동표", "현금흐름표"]:
            df = fs.get(name)
            if df is None or df.empty:
                continue
            df.to_excel(xw, sheet_name=name[:31], index=False)
    return out_path


# ---------- CLI ----------

def _process_one(pdf: Path, out_dir: Optional[Path] = None, verbose: bool = True) -> Optional[Path]:
    print(f"\n▶ {pdf.name}")
    try:
        fs = extract_financials(pdf, verbose=verbose)
    except Exception as e:
        print(f"  [에러] {e}")
        return None
    found = [n for n in ["재무상태표", "포괄손익계산서", "자본변동표", "현금흐름표"] if fs.get(n) is not None]
    print(f"  추출 섹션: {found}")
    for w in fs.get("_warnings", []):
        print(f"  [WARNING] {w}")
    if verbose:
        for n in found:
            df = fs[n]
            print(f"    - {n}: {df.shape[0]}행 × {df.shape[1]}열, 단위={df.attrs.get('unit')}, 페이지={df.attrs.get('pages')}")
    out = (out_dir or pdf.parent) / f"{pdf.stem}.xlsx"
    save_excel(fs, out)
    print(f"  ✓ 저장: {out}")
    return out


def main(argv: list[str]):
    if len(argv) < 2:
        print(__doc__)
        sys.exit(1)
    target = Path(argv[1])
    if target.is_file():
        _process_one(target)
    elif target.is_dir():
        pdfs = sorted(target.glob("*감사보고서*.pdf"))
        if not pdfs:
            pdfs = sorted(target.glob("*.pdf"))
        out_dir = target / "추출결과"
        out_dir.mkdir(exist_ok=True)
        results = []
        for pdf in pdfs:
            r = _process_one(pdf, out_dir=out_dir, verbose=False)
            if r:
                results.append(r)
        print(f"\n총 {len(results)}/{len(pdfs)} 처리 완료 → {out_dir}")
    else:
        print(f"경로를 찾을 수 없음: {target}")
        sys.exit(1)


if __name__ == "__main__":
    main(sys.argv)
