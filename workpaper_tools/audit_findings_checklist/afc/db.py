"""SQLite 저장소. 지적사항 DB + LLM 캐시 + 처리 로그."""
import json
import sqlite3
from datetime import datetime
from pathlib import Path

SCHEMA = """
CREATE TABLE IF NOT EXISTS source_files (
    file_hash TEXT PRIMARY KEY,
    file_name TEXT NOT NULL,
    issuer TEXT,
    doc_year INTEGER,
    n_pages INTEGER,
    extract_mode TEXT,          -- text | vision
    processed_at TEXT
);
CREATE TABLE IF NOT EXISTS pages (
    file_hash TEXT NOT NULL,
    page_no INTEGER NOT NULL,
    text TEXT,
    n_chars INTEGER,
    needs_ocr INTEGER DEFAULT 0,
    PRIMARY KEY (file_hash, page_no)
);
CREATE TABLE IF NOT EXISTS segments (
    segment_id TEXT PRIMARY KEY,   -- {파일해시8}-{순번}
    file_hash TEXT NOT NULL,
    seq INTEGER NOT NULL,
    title TEXT,
    page_start INTEGER,
    page_end INTEGER,
    pages_json TEXT,               -- [{"page": n, "text": "..."}]
    method TEXT,                   -- 규칙 | LLM
    profile TEXT
);
CREATE TABLE IF NOT EXISTS findings (
    finding_id TEXT PRIMARY KEY,   -- 공식 사례번호(FSS-2512-10, KICPA-2025-01) 또는 F-{연도}-{파일해시8}-{순번}
    segment_id TEXT NOT NULL,
    case_no TEXT,                  -- 원문에 적힌 공식 사례번호
    issue_area TEXT,               -- 원문 머리말의 쟁점 분야
    source_file TEXT,
    source_page INTEGER,
    source_page_end INTEGER,
    source_excerpt TEXT,
    issuer TEXT,
    year INTEGER,
    fiscal_period TEXT,
    title TEXT,
    finding_target TEXT,           -- 회사 | 감사인 | 회사 및 감사인
    standards TEXT,                -- JSON
    related_accounts TEXT,         -- JSON
    finding_type TEXT,
    risk_summary TEXT,
    audit_hint TEXT,               -- AI 제안(원문 아님)
    excerpt_status TEXT,           -- 일치 | 일치(문자만) | 유사 | 불일치 | 발췌없음
    excerpt_match_ratio REAL,
    standards_check TEXT,          -- JSON: 기준서별 원문 확인 결과
    segment_method TEXT,
    extract_mode TEXT,
    llm_model TEXT,
    prompt_version TEXT,
    llm_original_json TEXT,
    review_status TEXT DEFAULT '미검토',   -- 미검토 | 확정 | 제외
    reviewer_note TEXT,
    reviewed_at TEXT,
    created_at TEXT
);
CREATE TABLE IF NOT EXISTS llm_cache (
    cache_key TEXT PRIMARY KEY,
    purpose TEXT,
    model TEXT,
    response_text TEXT,
    input_tokens INTEGER,
    output_tokens INTEGER,
    created_at TEXT
);
CREATE TABLE IF NOT EXISTS process_log (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    ts TEXT,
    stage TEXT,
    level TEXT,                    -- INFO | WARNING | ERROR
    file_name TEXT,
    message TEXT
);
"""


def now() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


class Store:
    def __init__(self, path: Path):
        self.conn = sqlite3.connect(path)
        self.conn.row_factory = sqlite3.Row
        self.conn.executescript(SCHEMA)

    def close(self) -> None:
        self.conn.commit()
        self.conn.close()

    def log(self, stage: str, level: str, file_name: str, message: str) -> None:
        self.conn.execute(
            "INSERT INTO process_log (ts, stage, level, file_name, message) VALUES (?,?,?,?,?)",
            (now(), stage, level, file_name, message),
        )
        self.conn.commit()

    def upsert(self, table: str, row: dict) -> None:
        cols = ", ".join(row)
        marks = ", ".join("?" for _ in row)
        values = [json.dumps(v, ensure_ascii=False) if isinstance(v, (list, dict)) else v for v in row.values()]
        self.conn.execute(f"INSERT OR REPLACE INTO {table} ({cols}) VALUES ({marks})", values)

    def query(self, sql: str, params: tuple = ()) -> list[sqlite3.Row]:
        return self.conn.execute(sql, params).fetchall()

    def commit(self) -> None:
        self.conn.commit()
