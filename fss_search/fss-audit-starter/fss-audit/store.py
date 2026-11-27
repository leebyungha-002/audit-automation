"""SQLite 저장: 내용이 바뀌면 새 버전으로 추가(이력 보존)"""
import hashlib
import json
import sqlite3
import time
from datetime import date, datetime

import fss_client
from fss_client import EXTRA_SEARCH_FIELDS, clean_html


def connect(path: str = "fss.db") -> sqlite3.Connection:
    con = sqlite3.connect(path)
    con.execute("""CREATE TABLE IF NOT EXISTS docs(
        api TEXT, content_id TEXT, hash TEXT, subject TEXT, body TEXT,
        reg_date TEXT, url TEXT, raw TEXT, fetched_at TEXT,
        UNIQUE(api, content_id, hash))""")
    return con


def save(con, api: str, items: list[dict]) -> int:
    now = datetime.now().astimezone().isoformat(timespec="seconds")
    before = con.total_changes
    with con:
        for it in items:
            raw = json.dumps(it, ensure_ascii=False, sort_keys=True)
            body = it.get("contentKor") or it.get("contentsKor") or ""
            con.execute("INSERT OR IGNORE INTO docs VALUES (?,?,?,?,?,?,?,?,?)", (
                api, str(it["contentId"]), hashlib.sha256(raw.encode()).hexdigest(),
                clean_html(it.get("subject", "")), clean_html(body),
                it.get("regDate"), it.get("originUrl"), raw, now))
    return con.total_changes - before


def collect(con, key: str, first: int, last: int, apis=None, pause: float = 1.0):
    """API별·연도별 수집. (api, year, 수신건수, 신규/변경건수, 오류) 를 차례로 돌려준다.
    최대 조회기간이 미확인이라 연 단위로 나눠 호출하고,
    한 API에서 오류가 나면 그 API의 남은 연도만 건너뛴다(다른 API는 계속)."""
    today = date.today()
    for api in apis or fss_client.ENDPOINTS:
        for year in range(first, min(last, today.year) + 1):
            end = min(date(year, 12, 31), today).isoformat()
            try:
                items = fss_client.fetch(api, f"{year}-01-01", end, key)
            except fss_client.FssApiError as e:
                yield api, year, 0, 0, str(e)
                break
            yield api, year, len(items), save(con, api, items), None
            time.sleep(pause)


def _where(keyword: str, api: str | None) -> tuple[str, list]:
    """제목/본문/관련기준서·쟁점분야 검색 조건 (첨부파일 내용은 대상 아님)"""
    like = f"%{keyword}%"
    conds = ["subject LIKE ?", "body LIKE ?"]
    conds += [f"json_extract(raw, '$.{f}') LIKE ?" for f in EXTRA_SEARCH_FIELDS]
    sql = "(" + " OR ".join(conds) + ")"
    params = [like] * len(conds)
    if api:
        sql += " AND api = ?"
        params.append(api)
    return sql, params


def search(con, keyword: str, api: str | None = None) -> list[tuple]:
    where, params = _where(keyword, api)
    return con.execute(
        f"""SELECT api, reg_date, subject, url, fetched_at, body FROM docs
            WHERE {where} ORDER BY reg_date DESC""", params).fetchall()


def snippet(body: str, keyword: str, width: int = 60) -> str:
    """본문에서 검색어 주변만 잘라 보여준다 (없으면 앞부분)"""
    text = " ".join((body or "").split())
    i = text.find(keyword)
    if i < 0:
        return text[: width * 2] + ("…" if len(text) > width * 2 else "")
    start, stop = max(0, i - width), i + len(keyword) + width
    return ("…" if start else "") + text[start:stop] + ("…" if stop < len(text) else "")


def search_attachments(con, keyword: str, api: str | None = None) -> list[tuple[str, dict]]:
    """검색 결과 중 첨부파일이 있는 항목의 (API 구분, 원본 데이터)"""
    where, params = _where(keyword, api)
    rows = con.execute(
        f"""SELECT api, raw FROM docs
            WHERE {where} ORDER BY reg_date DESC, fetched_at DESC""", params).fetchall()
    found = [(api, json.loads(raw)) for api, raw in rows]
    return [(api, item) for api, item in found if item.get("atchfileUrl")]
