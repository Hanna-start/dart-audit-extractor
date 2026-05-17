# -*- coding: utf-8 -*-
"""
추출결과/ 폴더의 PDF별 xlsx 파일들을 모아서 시계열 통합 Excel 생성

- 별도/연결 분리
- 핵심 계정 (매출액/영업이익/당기순이익/자산총계 등) 자동 식별 → 연도×계정 매트릭스
- 원본 raw도 같이 합쳐서 검증 가능하게 보존
"""
from __future__ import annotations
import io
import re
import sys
from pathlib import Path

import pandas as pd

if __name__ == "__main__":
    for _s in (sys.stdout, sys.stderr):
        try:
            if hasattr(_s, "reconfigure"):
                _s.reconfigure(encoding="utf-8")
        except Exception:
            pass

# 독립 실행 시 사용하는 기본 경로 (라이브러리로 쓸 땐 run()에 직접 전달)
WORK = Path(__file__).resolve().parent
DEFAULT_SRC_DIR = WORK / "추출결과"
DEFAULT_OUT_FILE = WORK / "시계열_재무제표.xlsx"


# 핵심 계정 패턴 (공백 제거하고 매칭)
# 패턴은 _norm(공백 제거) 텍스트에 contains-매칭.
# `$` 끝 앵커 사용으로 "Ⅵ.당기순이익", "Ⅷ.당기순이익(손실)" 등 로마숫자 prefix 무관 매칭.
# transform: "abs" → 비용/지출 항목은 부호 무시하고 절댓값 사용 (괄호표기 양식 대응)
KEY_ITEMS = {
    "재무상태표": [
        ("유동자산",         r"유동자산$",          None),
        ("비유동자산",       r"비유동자산$",        None),
        ("자산총계",         r"자산총계$",          None),
        ("유동부채",         r"유동부채$",          None),
        ("비유동부채",       r"비유동부채$",        None),
        ("부채총계",         r"부채총계$",          None),
        ("자본금",           r"자본금$",            None),
        ("자본총계",         r"자본총계$",          None),
    ],
    "포괄손익계산서": [
        ("매출액",           r"매출액$|영업수익$",                              None),
        ("매출원가",         r"매출원가$",                                      "abs"),
        ("매출총이익",       r"매출총이익$",                                    None),
        ("판매비와관리비",   r"판매비와관리비$|판매비및관리비$",                "abs"),
        ("영업이익",         r"영업이익$|영업손실$|영업이익\(손실\)$|영업손익$", "loss_negate"),
        ("법인세차감전순이익", r"법인세(비용|등)?차감전(순이익|순손실|순이익\(손실\))$", "loss_negate"),
        ("법인세비용",       r"법인세비용$|^법인세$|법인세등$",                  "abs"),
        ("당기순이익",       r"당기순이익$|당기순손실$|당기순이익\(손실\)$",     "loss_negate"),
        ("총포괄이익",       r"총포괄이익$|총포괄손실$|총포괄이익\(손실\)$|총포괄손익$", "loss_negate"),
    ],
    "현금흐름표": [
        ("영업활동현금흐름", r"영업활동(으로?인한|에의한)?현금흐름$", None),
        ("투자활동현금흐름", r"투자활동(으로?인한|에의한)?현금흐름$", None),
        ("재무활동현금흐름", r"재무활동(으로?인한|에의한)?현금흐름$", None),
        ("기말현금",         r"기말(의)?현금(및현금성자산)?$",         None),
    ],
}


def _norm(s) -> str:
    return re.sub(r"\s+", "", str(s))


# 회사별 KEY_ITEMS 오버라이드 로드 ----------
# references/assumptions/<회사>/key_items.md 의 YAML front-matter를 읽는다.
# PyYAML 의존성 회피 — 간단한 줄 단위 파서.

OVERRIDE_ROOT = (
    Path(__file__).resolve().parent
    / ".claude" / "skills" / "dart-pipeline" / "references" / "assumptions"
)


def _parse_front_matter(text: str) -> dict | None:
    """YAML front-matter (--- ... ---) 한 블록을 단순 파싱.
    형식 가정: 키: 값, 들여쓰기 2/4칸, 리스트는 '- key: ...' / '  pattern: ...' / '  transform: ...'.
    실패하면 None.
    """
    lines = text.splitlines()
    if not lines or lines[0].strip() != "---":
        return None
    try:
        end_idx = next(i for i in range(1, len(lines)) if lines[i].strip() == "---")
    except StopIteration:
        return None

    overrides_by_section: dict[str, list[tuple]] = {}
    current_section: str | None = None
    current_entry: dict | None = None

    in_overrides = False
    for raw in lines[1:end_idx]:
        if raw.strip().startswith("#") or not raw.strip():
            continue
        if raw.startswith("overrides:"):
            in_overrides = True
            continue
        if not in_overrides:
            continue
        # 들여쓰기 깊이 추정
        stripped = raw.lstrip()
        indent = len(raw) - len(stripped)
        if indent == 2 and stripped.endswith(":"):
            # 재무제표 이름 (예: "  포괄손익계산서:")
            current_section = stripped.rstrip(":").strip()
            overrides_by_section.setdefault(current_section, [])
            current_entry = None
        elif stripped.startswith("- "):
            # 새 항목 시작 "- key: 매출액"
            current_entry = {}
            kv = stripped[2:].split(":", 1)
            if len(kv) == 2:
                current_entry[kv[0].strip()] = kv[1].strip().strip('"').strip("'")
            if current_section:
                overrides_by_section[current_section].append(current_entry)
        else:
            # 기존 항목에 속성 추가 "  pattern: ..." / "  transform: ..."
            if current_entry is None:
                continue
            kv = stripped.split(":", 1)
            if len(kv) == 2:
                val = kv[1].strip().strip('"').strip("'")
                if val.lower() == "null" or val == "":
                    val = None
                current_entry[kv[0].strip()] = val

    # 결과를 KEY_ITEMS 형식 [(key, pattern, transform), ...] 으로 변환
    out: dict[str, list[tuple]] = {}
    for sect, entries in overrides_by_section.items():
        result = []
        for e in entries:
            key = e.get("key")
            pat = e.get("pattern")
            tr = e.get("transform") or None
            if key and pat:
                result.append((key, pat, tr))
        if result:
            out[sect] = result
    return out


def load_override(company: str) -> dict | None:
    """회사명에 매칭되는 key_items.md 의 오버라이드 반환. 없으면 None."""
    if not company:
        return None
    f = OVERRIDE_ROOT / company / "key_items.md"
    if not f.exists():
        return None
    try:
        text = f.read_text(encoding="utf-8")
        parsed = _parse_front_matter(text)
        if parsed:
            print(f"   [override] {f.relative_to(Path(__file__).resolve().parent)} 적용 (섹션 {list(parsed.keys())})")
        return parsed
    except Exception as e:
        print(f"   [override 로드 실패] {f}: {e}")
        return None


_PREFIX_RE = re.compile(r"^[ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩⅪⅫIVXLM0-9\.\(\)]+")


def _norm_label(s) -> str:
    """첫 컬럼 라벨 정규화: 공백 제거 + 로마숫자/번호 prefix 제거 + 괄호 주석 제거
    예: 'Ⅵ. 당기순이익(손실)' → '당기순이익(손실)'
        '1. 매출액' → '매출액'
        '현금및현금성자산(주석2)' → '현금및현금성자산'
    """
    t = re.sub(r"\s+", "", str(s))
    t = _PREFIX_RE.sub("", t)
    # "(주석X)", "(주N)", "(주N,M)" 등 주석 표기 제거
    # 단 "(손실)" 같은 항목명 부분은 유지 → "주" + 숫자/콤마/공백만 허용
    t = re.sub(r"\(주(석)?[\d,\s]+\)", "", t)
    return t


def _split_period_columns(df: pd.DataFrame) -> tuple[list[str], list[str]]:
    """num 컬럼들을 당기/전기 그룹으로 분리.
    헤더에 '당'/'전' 표시가 있는 컬럼을 기준으로, 익명(col_N_num)은 직전 그룹에 귀속.
    """
    num_cols = [c for c in df.columns if str(c).endswith("_num")]
    cur, prev = [], []
    state = None
    for c in num_cols:
        s = str(c).replace("_num", "")
        if "당" in s:
            state = "cur"
            cur.append(c)
        elif "전" in s:
            state = "prev"
            prev.append(c)
        elif "주석" in s:
            # 주석 컬럼은 어느 그룹도 아님
            continue
        else:
            # 익명 컬럼 (col_2_num 등): 직전 state에 귀속
            if state == "cur":
                cur.append(c)
            elif state == "prev":
                prev.append(c)
    # fallback: 라벨이 전혀 없으면 합계 큰 두 컬럼을 당기/전기로
    if not cur and num_cols:
        ranked = sorted(
            [c for c in num_cols if "주석" not in str(c)],
            key=lambda c: df[c].abs().sum(),
            reverse=True,
        )
        cur = ranked[:max(1, len(ranked) // 2)]
        prev = ranked[len(cur):]
    return cur, prev


def _row_value(row: pd.Series, cols: list[str]) -> float | None:
    """행의 여러 _num 컬럼 중 0이 아닌 값 우선. 없으면 None."""
    vals = []
    for c in cols:
        v = row.get(c)
        if pd.notna(v):
            vals.append(v)
    nonzero = [v for v in vals if v != 0]
    if nonzero:
        # 같은 행에 0 아닌 값이 여러 개면 절댓값 최대 (보통 같은 값이 흩어진 경우)
        return max(nonzero, key=abs)
    return vals[0] if vals else None


def _loss_negate(val: float | None, matched_label: str) -> float | None:
    """K-GAAP 양식 대응 — 라벨이 "X손실"이면 양수로 적힌 값을 음수로 뒤집는다.

    예: "당기순손실 2,202,548,443" → -2,202,548,443
    단, "당기순이익(손실)" 처럼 부호 표기를 위해 (손실)이 붙은 경우는 그대로 둔다
    (이 경우 값에 이미 부호가 있음).
    """
    if val is None or pd.isna(val):
        return val
    # "(손실)" 형태 (괄호로 부호 표기 안내) → 값에 이미 부호 있으므로 그대로
    if "(손실)" in matched_label:
        return val
    # "손실"이 라벨에 있고 "이익"이 없으면 적자 라벨. 양수 값을 음수로 뒤집음.
    if "손실" in matched_label and "이익" not in matched_label:
        return -abs(val)
    return val


def _extract_key_values(df: pd.DataFrame, patterns: list) -> dict:
    """첫 컬럼에서 패턴 매칭 → {계정명: (당기값, 전기값)}
    patterns: [(key, regex_pattern, transform_or_None), ...]
        transform="abs":         비용·지출 → 절댓값
        transform="loss_negate": 손익 항목 → 라벨이 "X손실"이면 부호 반전 (K-GAAP 대응)
    """
    if df.empty:
        return {}
    cur_cols, prev_cols = _split_period_columns(df)
    if not cur_cols:
        return {}

    first_col = df.columns[0]
    norm_first = df[first_col].apply(_norm_label)

    out = {}
    for entry in patterns:
        key, pat = entry[0], entry[1]
        transform = entry[2] if len(entry) > 2 else None
        # _norm_label로 정규화된 라벨에 fullmatch (정확 일치)
        rx = re.compile(_norm(pat).rstrip("$") + "$")  # 패턴 끝 $ 보장
        hits_mask = norm_first.str.fullmatch(rx, na=False)
        hits = df[hits_mask]
        if hits.empty:
            continue
        best_idx = None
        best_val = None
        for idx, row in hits.iterrows():
            v = _row_value(row, cur_cols)
            if v is None or v == 0:
                continue
            if best_val is None or abs(v) > abs(best_val):
                best_idx, best_val = idx, v
        if best_idx is None:
            best_idx = hits.index[0]
            best_val = _row_value(hits.loc[best_idx], cur_cols)
        prev_val = _row_value(df.loc[best_idx], prev_cols) if prev_cols else None

        if transform == "abs":
            if best_val is not None and pd.notna(best_val):
                best_val = abs(best_val)
            if prev_val is not None and pd.notna(prev_val):
                prev_val = abs(prev_val)
        elif transform == "loss_negate":
            matched_label = norm_first.loc[best_idx]
            best_val = _loss_negate(best_val, matched_label)
            prev_val = _loss_negate(prev_val, matched_label)

        out[key] = (best_val, prev_val)
    return out


def run(src_dir: Path | str, out_file: Path | str, default_company: str = "회사") -> Path:
    """추출결과 폴더의 PDF별 xlsx → 시계열 통합 Excel.

    Args:
        src_dir: PDF별 변환 결과 xlsx 폴더
        out_file: 통합 결과 Excel 경로
        default_company: 메타에 회사명이 비어있을 때 사용할 폴백명
    Returns:
        저장된 out_file 경로
    """
    src_dir = Path(src_dir)
    out_file = Path(out_file)
    files = sorted(src_dir.glob("*.xlsx"))
    print(f"입력 파일: {len(files)}개")

    # 파일별로 메타 + 핵심 계정 수집
    records = []  # [{회사명, 사업연도, 구분, 재무제표, 계정, 값}, ...]
    raw_collected = {sj: [] for sj in ["재무상태표", "포괄손익계산서", "자본변동표", "현금흐름표"]}

    for f in files:
        try:
            meta_df = pd.read_excel(f, sheet_name="meta")
        except Exception as e:
            print(f"  [skip] {f.name}: {e}")
            continue
        meta = dict(zip(meta_df["항목"], meta_df["값"]))
        company = meta.get("회사명") or default_company
        year = int(meta.get("사업연도")) if pd.notna(meta.get("사업연도")) else None
        prev_year = int(meta.get("전기연도")) if pd.notna(meta.get("전기연도")) else None
        div = meta.get("구분") or "별도"

        print(f"\n[{f.name}]  {company} / {year} / {div}")

        # 회사별 KEY_ITEMS 오버라이드 적용
        # (이 파일의 회사명 기준 — 같은 회사명이면 같은 오버라이드)
        override = load_override(company)
        effective_key_items = dict(KEY_ITEMS)
        if override:
            for sect, items in override.items():
                if sect in effective_key_items:
                    effective_key_items[sect] = items

        for sj, patterns in effective_key_items.items():
            try:
                df = pd.read_excel(f, sheet_name=sj)
            except Exception:
                continue
            # raw 보존 (메타 컬럼 부착)
            raw = df.copy()
            raw.insert(0, "구분", div)
            raw.insert(0, "사업연도", year)
            raw_collected[sj].append(raw)

            extracted = _extract_key_values(df, patterns)
            for key, (cur, prev) in extracted.items():
                if pd.notna(cur):
                    records.append({
                        "회사명": company, "사업연도": year, "구분": div,
                        "재무제표": sj, "계정": key, "값": cur,
                    })
                if prev is not None and pd.notna(prev) and prev_year:
                    records.append({
                        "회사명": company, "사업연도": prev_year, "구분": div,
                        "재무제표": sj, "계정": key, "값": prev,
                    })
            print(f"   {sj}: {len(extracted)}개 계정 추출")

    if not records:
        print("[!] 수집된 데이터 없음")
        return out_file

    long_df = pd.DataFrame(records)

    # 같은 (회사,연도,구분,재무제표,계정) 조합이 여러 PDF에서 나오면 (당기/전기 중복)
    # 더 최근 보고서 = 더 정확한 값 → 그냥 last 또는 mean (보통 동일)
    long_df = long_df.groupby(["회사명", "사업연도", "구분", "재무제표", "계정"], as_index=False).agg({"값": "last"})

    # 시트별 피벗 (연도 × 계정)
    out_file.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(out_file, engine="openpyxl") as xw:
        # 1) 통합 long-format
        long_df.to_excel(xw, sheet_name="long_data", index=False)

        # 2) 핵심 시계열 시트 (구분별, 재무제표별 피벗)
        for div in long_df["구분"].unique():
            for sj in long_df["재무제표"].unique():
                sub = long_df[(long_df["구분"] == div) & (long_df["재무제표"] == sj)]
                if sub.empty:
                    continue
                # 계정 순서를 KEY_ITEMS 정의 순서로
                order = [entry[0] for entry in KEY_ITEMS[sj]]
                pv = sub.pivot_table(index="계정", columns="사업연도", values="값", aggfunc="last")
                pv = pv.reindex([k for k in order if k in pv.index])
                sheet = f"{div}_{sj}"[:31]
                pv.to_excel(xw, sheet_name=sheet)

        # 3) raw 시트 (PDF별 원본 통합)
        for sj, dfs in raw_collected.items():
            if dfs:
                combined = pd.concat(dfs, ignore_index=True, sort=False)
                combined.to_excel(xw, sheet_name=f"raw_{sj}"[:31], index=False)

    print(f"\n✓ 저장 완료: {out_file}")
    print(f"  - long_data: {len(long_df)}행")
    print(f"  - 시계열 피벗 시트: 구분({long_df['구분'].nunique()}) × 재무제표(4)")
    return out_file


if __name__ == "__main__":
    run(DEFAULT_SRC_DIR, DEFAULT_OUT_FILE)
