# -*- coding: utf-8 -*-
"""① 상장사 모집단(universe) 확정.

CORPCODE.xml(전체 ~12만사)에서 stock_code가 있는 상장사만 추려
corp_index.json 으로 캐시한다. (개선안 D 실현 — 전체 스크리닝의 전제)

사용:
  py screener/universe.py                 # 인덱스 빌드 + 통계
  py screener/universe.py --pilot 20      # 파일럿 표본 corp_code 출력
"""
from __future__ import annotations

import argparse
import json
import sys
import xml.etree.ElementTree as ET
from datetime import date
from pathlib import Path

if __name__ == "__main__":
    for _s in (sys.stdout, sys.stderr):
        try:
            if hasattr(_s, "reconfigure"):
                _s.reconfigure(encoding="utf-8")
        except Exception:
            pass

from common import CORPCODE_XML, CORP_INDEX, get_conn


def build_index(corpcode_xml: Path = CORPCODE_XML, out: Path = CORP_INDEX) -> list[dict]:
    """상장사(stock_code 보유)만 추려 corp_index.json 저장."""
    if not corpcode_xml.exists():
        raise FileNotFoundError(
            f"{corpcode_xml} 없음. 먼저 dart_client.ensure_corpcode 로 받으세요.")
    tree = ET.parse(str(corpcode_xml))
    listed = []
    for item in tree.getroot().iter("list"):
        stock = (item.findtext("stock_code") or "").strip()
        if not stock:                 # 비상장 제외
            continue
        listed.append({
            "corp_code": (item.findtext("corp_code") or "").strip(),
            "corp_name": (item.findtext("corp_name") or "").strip(),
            "stock_code": stock,
            "modify_date": (item.findtext("modify_date") or "").strip(),
        })
    listed.sort(key=lambda c: c["corp_code"])
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(listed, ensure_ascii=False, indent=2), encoding="utf-8")

    # companies 테이블에도 기본 행 upsert (sector는 collect 단계에서 채움)
    conn = get_conn()
    today = date.today().isoformat()
    with conn:
        for c in listed:
            conn.execute(
                "INSERT INTO companies(corp_code, corp_name, stock_code, updated_at) "
                "VALUES(?,?,?,?) ON CONFLICT(corp_code) DO UPDATE SET "
                "corp_name=excluded.corp_name, stock_code=excluded.stock_code",
                (c["corp_code"], c["corp_name"], c["stock_code"], today),
            )
    conn.close()
    return listed


def load_index(path: Path = CORP_INDEX) -> list[dict]:
    if not path.exists():
        return build_index()
    return json.loads(path.read_text(encoding="utf-8"))


def pilot_sample(n: int = 20, path: Path = CORP_INDEX) -> list[dict]:
    """모집단 전체에 고르게 분포된 결정적 표본 n개.

    corp_code 정렬 후 등간격 추출 → 시장·업종이 한쪽에 쏠리지 않게
    (금융/비IFRS 같은 엣지케이스가 표본에 섞일 확률을 높임).
    """
    idx = load_index(path)
    if not idx:
        return []
    if n >= len(idx):
        return idx
    step = len(idx) / n
    return [idx[int(i * step)] for i in range(n)]


def main(argv):
    ap = argparse.ArgumentParser()
    ap.add_argument("--pilot", type=int, default=0, help="파일럿 표본 크기 출력")
    args = ap.parse_args(argv[1:])

    listed = build_index()
    print(f"상장사 모집단: {len(listed)}개사 → {CORP_INDEX.relative_to(CORPCODE_XML.parent.parent)}")

    if args.pilot:
        sample = pilot_sample(args.pilot)
        print(f"\n파일럿 표본 {len(sample)}개 (등간격):")
        for c in sample:
            print(f"  {c['corp_code']}  {c['stock_code']}  {c['corp_name']}")


if __name__ == "__main__":
    main(sys.argv)
