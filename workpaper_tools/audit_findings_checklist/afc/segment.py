"""지적사항 단위 분할. 규칙(segment_rules.yaml) 우선, 맞는 규칙이 없으면 LLM 보조."""
import re

from .db import Store
from .llm import LLM, LLMError

SEGMENT_SYSTEM = """당신은 한국의 회계 심사·감리 지적사례 문서를 개별 지적사례 단위로 나누는 보조자입니다.
입력은 [p.N] 표시로 페이지가 구분된 문서 텍스트입니다.
표지, 목차, 통계, 개요, 안내 페이지는 제외하고, 개별 지적사례(특정 회사나 감사인에 대한 지적 1건)만 골라
사례별 제목과 시작·끝 페이지를 반환하세요. 제목은 문서에 적힌 표현을 그대로 사용합니다.
개별 지적사례가 하나도 없으면 빈 목록을 반환하세요."""

SEGMENT_SCHEMA = {
    "type": "object",
    "properties": {
        "segments": {
            "type": "array",
            "items": {
                "type": "object",
                "properties": {
                    "title": {"type": "string"},
                    "start_page": {"type": "integer"},
                    "end_page": {"type": "integer"},
                },
                "required": ["title", "start_page", "end_page"],
                "additionalProperties": False,
            },
        }
    },
    "required": ["segments"],
    "additionalProperties": False,
}


def _pick_profile(pages: list[dict], profiles: list[dict]) -> dict | None:
    for prof in profiles:
        pat = re.compile(prof["pattern"], re.MULTILINE)
        if prof["type"] == "page_title":
            hits = sum(1 for p in pages if pat.search(p["text"]))
        else:
            hits = sum(len(pat.findall(p["text"])) for p in pages)
        if hits >= prof["min_hits"]:
            return prof
    return None


def _split_page_title(pages: list[dict], prof: dict) -> list[dict]:
    pat = re.compile(prof["pattern"], re.MULTILINE)
    segments: list[dict] = []
    current_key = None
    for p in pages:
        keys = {re.sub(r"\s+", "", m.group("key")) for m in pat.finditer(p["text"])}
        if len(keys) != 1:  # 0개: 사례 외 페이지, 2개 이상: 목차
            current_key = None
            continue
        key = keys.pop()
        if key != current_key:
            title = next(m.group(0) for m in pat.finditer(p["text"]))
            segments.append({"title": title.strip(), "pages": []})
            current_key = key
        segments[-1]["pages"].append({"page": p["page"], "text": p["text"]})
    return segments


def _split_anchor(pages: list[dict], prof: dict) -> list[dict]:
    pat = re.compile(prof["pattern"])
    lines = [(p["page"], ln) for p in pages for ln in p["text"].splitlines()]
    starts: list[tuple[int, str]] = []  # (구간 시작 줄 인덱스, 제목)
    for i, (_, ln) in enumerate(lines):
        if not pat.match(ln.strip()):
            continue
        if prof["title_mode"] == "anchor_plus_next":
            extra = [lines[j][1].strip() for j in range(i + 1, min(i + 1 + prof.get("title_lines", 1), len(lines)))]
            starts.append((i, " ".join([ln.strip(), *extra])))
        else:  # after_number_line
            begin, title = i, ""
            for back in range(1, prof.get("title_lookback", 4) + 1):
                if i - back < 0:
                    break
                if re.fullmatch(r"\d{1,2}", lines[i - back][1].strip()):
                    begin = i - back
                    title = " ".join(lines[j][1].strip() for j in range(begin + 1, i))
                    break
            if not title and i > 0:
                begin, title = i - 1, lines[i - 1][1].strip()
            starts.append((begin, title))

    segments = []
    for n, (begin, title) in enumerate(starts):
        end = starts[n + 1][0] if n + 1 < len(starts) else len(lines)
        by_page: dict[int, list[str]] = {}
        for page, ln in lines[begin:end]:
            by_page.setdefault(page, []).append(ln)
        segments.append({
            "title": title,
            "pages": [{"page": pg, "text": "\n".join(lns)} for pg, lns in by_page.items()],
        })
    return segments


def _split_with_llm(pages: list[dict], llm: LLM) -> list[dict]:
    text = "\n\n".join(f"[p.{p['page']}]\n{p['text']}" for p in pages)
    data = llm.call_json("segment", llm.cfg["structure_model"], SEGMENT_SYSTEM,
                         [{"type": "text", "text": text}], SEGMENT_SCHEMA)
    by_no = {p["page"]: p for p in pages}
    segments = []
    for s in data["segments"]:
        span = [by_no[n] for n in range(s["start_page"], s["end_page"] + 1) if n in by_no]
        if span:
            segments.append({"title": s["title"], "pages": [{"page": p["page"], "text": p["text"]} for p in span]})
    return segments


def run_segment(cfg: dict, store: Store, log, force: bool = False, name_filter: str | None = None,
                force_llm: bool = False) -> None:
    llm = None
    for f in store.query("SELECT * FROM source_files ORDER BY file_name"):
        if name_filter and name_filter not in f["file_name"]:
            continue
        fhash = f["file_hash"]
        if not force and store.query("SELECT 1 FROM segments WHERE file_hash=?", (fhash,)):
            log.info("건너뜀(이미 분할됨): %s", f["file_name"])
            continue
        pages = [{"page": r["page_no"], "text": r["text"] or ""}
                 for r in store.query("SELECT page_no, text FROM pages WHERE file_hash=? ORDER BY page_no", (fhash,))]

        prof = None if force_llm else _pick_profile(pages, cfg["segment_profiles"])
        if prof:
            method, profile_name = "규칙", prof["name"]
            segments = _split_page_title(pages, prof) if prof["type"] == "page_title" else _split_anchor(pages, prof)
        else:
            method, profile_name = "LLM", None
            log.warning("%s: 맞는 분할 규칙 없음 → LLM 보조 분할", f["file_name"])
            try:
                llm = llm or LLM(cfg, store, log)
                segments = _split_with_llm(pages, llm)
            except (LLMError, ValueError, KeyError) as e:
                log.error("%s: LLM 분할 실패 (%s)", f["file_name"], e)
                store.log("segment", "ERROR", f["file_name"], f"LLM 분할 실패: {e}")
                continue

        store.conn.execute("DELETE FROM segments WHERE file_hash=?", (fhash,))
        for seq, seg in enumerate(segments, start=1):
            store.upsert("segments", {
                "segment_id": f"{fhash[:8]}-{seq:03d}", "file_hash": fhash, "seq": seq,
                "title": seg["title"][:200],
                "page_start": seg["pages"][0]["page"], "page_end": seg["pages"][-1]["page"],
                "pages_json": seg["pages"], "method": method, "profile": profile_name,
            })
        store.commit()
        level = "INFO" if segments else "WARNING"
        msg = f"{len(segments)}건 분할 (방식={method}, 규칙={profile_name})"
        log.log(20 if segments else 30, "%s: %s", f["file_name"], msg)
        store.log("segment", level, f["file_name"], msg)
    if llm:
        log.info(llm.usage_summary())
