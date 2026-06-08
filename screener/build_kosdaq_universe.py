# -*- coding: utf-8 -*-
"""코스닥 정식 명단 생성 (KRX/FinanceDataReader 기준, 권위 있는 현재 상장 목록).

build_kospi_universe.py의 코스닥판. FinanceDataReader('KOSDAQ')의 현재 상장 종목에서
우선주를 제외하고 DART corp_code로 매핑되는 '고유 회사'만 추려 kosdaq_index.json에 저장한다.
시가총액 큰 순 정렬(배치 수집 시 대형주 우선).

- 상폐 섞임 없음(현재 상장만), DART 전수 태깅 불필요.
- 주가/거래대금/시총은 FinanceDataReader가 함께 제공(수집·스크리닝 단계에서 활용).

사용: py screener/build_kosdaq_universe.py
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

if __name__ == "__main__":
    for _s in (sys.stdout, sys.stderr):
        try:
            if hasattr(_s, "reconfigure"):
                _s.reconfigure(encoding="utf-8")
        except Exception:
            pass

from common import DATA_DIR, CORP_INDEX

KOSDAQ_INDEX = DATA_DIR / "kosdaq_index.json"


def main():
    import FinanceDataReader as fdr

    df = fdr.StockListing("KOSDAQ")
    print(f"FDR 코스닥 종목(우선주 포함): {len(df)}")

    idx = json.loads(CORP_INDEX.read_text(encoding="utf-8"))
    s2c = {str(c["stock_code"]).strip().zfill(6): c for c in idx if c.get("stock_code")}

    # 시총 큰 순 정렬 후 corp_code 매핑되는 보통주만
    if "Marcap" in df.columns:
        df = df.sort_values("Marcap", ascending=False)

    rows, seen = [], set()
    for _, r in df.iterrows():
        sc = str(r["Code"]).strip().zfill(6)
        meta = s2c.get(sc)
        if not meta:
            continue                      # 우선주/전환주 등 corp_code 없음 → 제외
        cc = meta["corp_code"]
        if cc in seen:
            continue
        seen.add(cc)
        rows.append({
            "corp_code": cc,
            "corp_name": meta.get("corp_name", "") or str(r.get("Name", "")),
            "stock_code": sc,
            "marcap": int(r["Marcap"]) if "Marcap" in df.columns and r["Marcap"] == r["Marcap"] else None,
        })

    KOSDAQ_INDEX.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"코스닥 정식 명단: {len(rows)}개 회사 → {KOSDAQ_INDEX.name} (시총 큰 순)")
    print("상위 8:", ", ".join(f"{x['corp_name']}({x['stock_code']})" for x in rows[:8]))


if __name__ == "__main__":
    main()
