# -*- coding: utf-8 -*-
"""감사보고서 document.xml(XML) → 재무제표 핵심 계정 추출.

비상장·사업보고서 미제출 시기(정형 API 미커버)의 재무제표를, PDF 스크래핑 대신
공시 원문 XML의 표에서 직접 파싱한다(레이아웃 깨짐에 더 강함).

계정 매핑·당기전기 분리·손익 부호 처리는 consolidate.py 의 로직을 재사용한다(재구현 금지).

테스트:
  py dart_xml_fs.py 20240409000213 별도      # 무신사 2023 별도 감사보고서
"""
from __future__ import annotations

import io
import re
import sys
import zipfile
from pathlib import Path

import pandas as pd
import requests

import consolidate  # KEY_ITEMS, _extract_key_values 재사용

HERE = Path(__file__).resolve().parent
OPEN = "https://opendart.fss.or.kr/api"


def _load_key() -> str:
    import os
    k = os.environ.get("DART_API_KEY", "")
    if k:
        return k
    envp = HERE / ".env"
    if envp.exists():
        for line in envp.read_text(encoding="utf-8").splitlines():
            if line.strip().startswith("DART_API_KEY="):
                return line.split("=", 1)[1].strip().strip('"').strip("'")
    return ""


def _num(s):
    """'1,234' / '(1,234)' / '△1,234' / '-' → float|None."""
    t = re.sub(r"\s", "", str(s))
    if not t or t in ("-", "–", "—"):
        return None
    neg = False
    if t.startswith("(") and t.endswith(")"):
        neg = True; t = t[1:-1]
    if t[:1] in ("△", "▲", "-", "−"):
        neg = True; t = t[1:]
    t = t.replace(",", "")
    if not re.fullmatch(r"\d+(\.\d+)?", t):
        return None
    v = float(t)
    return -v if neg else v


def _cells(tr_html):
    return [re.sub(r"\s+", " ", re.sub(r"<[^>]+>", "", c)).strip()
            for c in re.findall(r"<T[HDE]\b[^>]*>(.*?)</T[HDE]>", tr_html, re.S | re.I)]


def _decode(b: bytes) -> str:
    """DART 문서 인코딩 자동 감지 — 신형 utf-8 / 구형 euc-kr(cp949)."""
    for enc in ("utf-8", "cp949", "euc-kr"):
        try:
            t = b.decode(enc)
            # 한글 마커가 보이면 그 인코딩 채택
            if any(k in t for k in ("재무제표", "감사보고서", "주식회사", "재무상태표", "대차대조표")):
                return t
        except (UnicodeDecodeError, LookupError):
            continue
    return b.decode("utf-8", "ignore")


def download_xml(key, rcept_no) -> str:
    """document.xml zip → 재무제표가 들어있는 멤버 텍스트(없으면 첫 멤버)."""
    r = requests.get(f"{OPEN}/document.xml",
                     params={"crtfc_key": key, "rcept_no": rcept_no}, timeout=120)
    z = zipfile.ZipFile(io.BytesIO(r.content))
    names = z.namelist()
    decoded = [_decode(z.read(n)) for n in names]
    for t in decoded:
        flat_ns = re.sub(r"\s", "", re.sub(r"<[^>]+>", "", t))
        if "자산총계" in flat_ns or "대차대조표" in flat_ns:
            return t
    return decoded[0] if decoded else ""


def _attached_segment(raw: str) -> str:
    """(첨부)재무제표 ~ 주석 사이로 한정 (없으면 전체)."""
    titles = [(m.start(), re.sub(r"\s", "", re.sub(r"<[^>]+>", "", m.group(1))))
              for m in re.finditer(r"<TITLE[^>]*>(.*?)</TITLE>", raw, re.S)]
    att = [p for p, t in titles if "첨부" in t and "재무제표" in t]
    note = [p for p, t in titles if t.startswith("주석")]
    if att:
        end = next((p for p, t in titles if t.startswith("주석") and p > att[0]), len(raw))
        return raw[att[0]:end]
    return raw


def _unit_scale(prefix_text: str) -> int:
    m = list(re.finditer(r"단위\s*:?\s*([백천]?)\s*원", prefix_text))
    if not m:
        return 1
    u = m[-1].group(1)
    return {"백": 1_000_000, "천": 1_000}.get(u, 1)


def _classify(flat: str) -> str | None:
    f = re.sub(r"\s", "", flat)   # 글자 사이 공백 제거 ('자 산 총 계' → '자산총계')
    has = lambda *k: any(x in f for x in k)
    if has("자산총계"):
        return "재무상태표"
    if has("영업활동") and "현금흐름" in f:
        return "현금흐름표"
    if has("매출총이익", "매출액", "영업수익", "영업이익", "영업손실"):
        return "포괄손익계산서"
    return None


def _table_to_df(tbl_html: str, scale: int) -> pd.DataFrame:
    rows = [_cells(r) for r in re.findall(r"<TR\b[^>]*>(.*?)</TR>", tbl_html, re.S | re.I)]
    rows = [r for r in rows if r]
    if not rows:
        return pd.DataFrame()
    # 헤더 행: 첫 셀이 '과목' 류이고 '기'/'당'/'전' 포함
    hdr_idx = 0
    for i, r in enumerate(rows[:5]):
        first = re.sub(r"\s", "", r[0])
        if first.startswith("과목") or first in ("과목", "계정과목", "구분"):
            hdr_idx = i; break
    ncol = max(len(r) for r in rows)
    header = rows[hdr_idx] + [""] * (ncol - len(rows[hdr_idx]))
    names = []
    for i, h in enumerate(header):
        h = h.strip()
        names.append(h if h else f"col_{i}")
    # 중복 이름 방지
    seen = {}
    for i, n in enumerate(names):
        if n in seen:
            seen[n] += 1; names[i] = f"{n}_{seen[n]}"
        else:
            seen[n] = 0
    data = []
    for r in rows[hdr_idx + 1:]:
        r = r + [""] * (ncol - len(r))
        data.append(r[:ncol])
    df = pd.DataFrame(data, columns=names)
    # _num 컬럼 부착 (첫 컬럼=과목 제외)
    for i, c in enumerate(df.columns):
        if i == 0:
            continue
        nums = df[c].map(_num)
        if nums.notna().sum() >= 1:
            df[f"{c}_num"] = nums * scale
    return df


def extract(key: str, rcept_no: str, kind: str, company: str,
            year: int, prev_year: int | None) -> list[dict]:
    """감사보고서 XML → long_data 레코드 리스트."""
    raw = download_xml(key, rcept_no)
    seg = _attached_segment(raw)
    # 표 + 직전 텍스트(단위) 위치
    recs = []
    used = set()
    for m in re.finditer(r"<TABLE\b.*?</TABLE>", seg, re.S | re.I):
        tbl = m.group(0)
        flat = re.sub(r"<[^>]+>", "", tbl)
        sj = _classify(flat)
        if sj is None or sj in used:
            continue
        # 본 재무제표만: 헤더에 당/전 기 표시가 있어야 함 (주석 재작성표 배제)
        f_ns = re.sub(r"\s", "", flat)
        if not re.search(r"\(당\)|당기|제\d+\(당", f_ns):
            continue
        scale = _unit_scale(seg[max(0, m.start() - 400):m.start()])
        df = _table_to_df(tbl, scale)
        if df.empty:
            continue
        extracted = consolidate._extract_key_values(df, consolidate.KEY_ITEMS[sj])
        if not extracted:
            continue
        used.add(sj)
        for keyname, (cur, prev) in extracted.items():
            if cur is not None and pd.notna(cur):
                recs.append({"회사명": company, "사업연도": year, "구분": kind,
                             "재무제표": sj, "계정": keyname, "값": int(round(cur))})
            if prev is not None and pd.notna(prev) and prev_year:
                recs.append({"회사명": company, "사업연도": prev_year, "구분": kind,
                             "재무제표": sj, "계정": keyname, "값": int(round(prev))})
    return recs


if __name__ == "__main__":
    key = _load_key()
    rcept = sys.argv[1] if len(sys.argv) > 1 else "20240409000213"
    kind = sys.argv[2] if len(sys.argv) > 2 else "별도"
    year = int(sys.argv[3]) if len(sys.argv) > 3 else 2023
    recs = extract(key, rcept, kind, "무신사", year, year - 1)
    df = pd.DataFrame(recs)
    for sj in ("재무상태표", "포괄손익계산서", "현금흐름표"):
        sub = df[(df["재무제표"] == sj)]
        print(f"\n== {sj} ==")
        for _, r in sub.sort_values("사업연도").iterrows():
            print(f"  {r['사업연도']} {r['계정']:14} {r['값']:>20,}")
