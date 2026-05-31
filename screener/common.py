# -*- coding: utf-8 -*-
"""스크리너 축 공통 — 경로·API키·sqlite 스키마.

축 1(회사별 심층 파이프라인)과 격리된 두 번째 축의 공유 기반.
헌법(결정성 분리)은 그대로 적용: 비율·진단 계산은 결정적 스크립트에서만.
"""
from __future__ import annotations

import os
import sqlite3
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent      # 프로젝트 루트
DATA_DIR = ROOT / "data"
CORPCODE_XML = DATA_DIR / "CORPCODE.xml"
CORP_INDEX = DATA_DIR / "corp_index.json"           # 개선안 D: 상장사 인덱스 캐시
DB_PATH = ROOT / "screener.db"                       # gitignore 대상

# api_to_timeseries.fetch_year / ID_MAP 재사용을 위해 루트를 import 경로에 추가
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))


def load_key(cli: str | None = None) -> str:
    """DART API 키 로드: --key > 환경변수 > .env (다른 모듈과 동일 규약)."""
    if cli:
        return cli.strip()
    k = os.environ.get("DART_API_KEY", "").strip()
    if k:
        return k
    envp = ROOT / ".env"
    if envp.exists():
        for line in envp.read_text(encoding="utf-8").splitlines():
            if line.strip().startswith("DART_API_KEY="):
                return line.split("=", 1)[1].strip().strip('"').strip("'")
    return ""


SCHEMA = """
CREATE TABLE IF NOT EXISTS companies (
    corp_code    TEXT PRIMARY KEY,
    corp_name    TEXT,
    stock_code   TEXT,
    corp_cls     TEXT,   -- Y=유가/K=코스닥/N=코넥스/E=기타 (company.json)
    induty_code  TEXT,   -- KSIC 업종코드 (64~66 = 금융·보험)
    is_financial INTEGER DEFAULT 0,
    updated_at   TEXT
);

CREATE TABLE IF NOT EXISTS financials (
    corp_code  TEXT,
    bsns_year  INTEGER,
    fs_div     TEXT,     -- OFS=별도 / CFS=연결
    재무제표    TEXT,
    계정        TEXT,
    값          INTEGER,
    PRIMARY KEY (corp_code, bsns_year, fs_div, 재무제표, 계정)
);

CREATE TABLE IF NOT EXISTS collect_log (
    corp_code   TEXT,
    bsns_year   INTEGER,
    fs_div      TEXT,
    status      TEXT,    -- ok / empty / error / cap
    n_accounts  INTEGER,
    fetched_at  TEXT,
    PRIMARY KEY (corp_code, bsns_year, fs_div)
);
"""


def get_conn(db_path: Path | str = DB_PATH) -> sqlite3.Connection:
    """sqlite 연결 + 스키마 보장. 재실행 안전."""
    conn = sqlite3.connect(str(db_path))
    conn.executescript(SCHEMA)
    return conn
