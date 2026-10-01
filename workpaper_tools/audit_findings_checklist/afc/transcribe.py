"""텍스트 층을 읽을 수 없는 PDF(글꼴 인코딩 깨짐, 스캔본)를 Claude가 페이지 이미지로 판독한다.

판독문은 pages 테이블에 들어가 이후 분할·구조화의 원문 노릇을 한다. 다만 이것은 PDF에서 직접 뽑은
텍스트가 아니라 AI 판독문이므로, 이 문서에서 나온 지적사항의 발췌 검증은 "일치(AI 판독문)"로 따로 표시된다.
"""
import base64
import json
import re
from concurrent.futures import ThreadPoolExecutor, as_completed

from .db import Store
from .extract import clean_page_text, render_page_png
from .llm import LLM, LLMError

TRANSCRIBE_SYSTEM = """당신은 한국어 회계 문서의 페이지 이미지를 텍스트로 옮기는 판독자입니다.
이 판독문은 뒤 단계에서 원문 노릇을 하며, 지적사례의 발췌가 실제로 문서에 있는지 글자 단위로 대조하는 기준이 됩니다. 그래서 요약하거나 고쳐 쓰지 말고, 이미지에 보이는 글자를 그대로 옮기는 것이 가장 중요합니다. 숫자·기준서 번호·문단 번호·연도는 특히 정확히 옮깁니다.

옮기는 방법:
- 본문 문단은 줄바꿈 없이 한 문단을 한 줄로 이어 쓰고, 문단 사이는 줄을 바꿉니다. 줄 끝에서 끊긴 단어는 이어 붙입니다.
- 제목과 머리말은 각각 한 줄로 씁니다. 사례 제목 줄은 사례번호로 시작하게 씁니다(예: "FSS/2106-01 : 매출 및 매출원가 허위계상"). 제목 앞의 순번 도형(①, 1 등)은 옮기지 않습니다.
- "▣ 쟁점 분야:", "▣ 관련 기준:", "▣ 결정일:", "▣ 회계결산일:" 같은 머리말 항목과 "1. 회사의 회계처리" 같은 소제목은 보이는 그대로 한 줄씩 씁니다. 소제목의 번호가 도형 안에 있어도 "1. 회사의 회계처리"처럼 번호를 붙여 씁니다.
- 페이지 위쪽에 반복되는 책 제목과 아래쪽 쪽번호는 옮기지 않습니다.
- 표는 행마다 한 줄, 칸은 " | "로 구분합니다. 도식·그림은 그 안의 글자를 옮기지 말고 "[도식: 제목]" 한 줄로 표시합니다.
- 글자가 없는 페이지(표지, 간지, 빈 면)는 text를 빈 문자열로 둡니다.
- 읽을 수 없는 글자는 추측하지 말고 □로 표시합니다."""

TRANSCRIBE_SCHEMA = {
    "type": "object",
    "properties": {"text": {"type": "string"}},
    "required": ["text"],
    "additionalProperties": False,
}


def run_transcribe(cfg: dict, store: Store, log, name_filter: str | None = None, limit: int | None = None) -> None:
    files = [f for f in store.query("SELECT * FROM source_files WHERE extract_mode='unreadable' ORDER BY file_name")
             if not name_filter or name_filter in f["file_name"]]
    if not files:
        log.info("판독이 필요한 파일이 없습니다.")
        return
    llm = LLM(cfg, store, log)
    tcfg = cfg["transcribe"]
    noise = [re.compile(p) for p in cfg["extract"]["noise_lines"]]

    for f in files:
        pdf_path = cfg["paths"]["input_pdfs"] / f["file_name"]
        page_nos = list(range(1, f["n_pages"] + 1))[:limit]
        log.info("%s: %d쪽 판독 시작 (모델 %s)", f["file_name"], len(page_nos), tcfg["model"])

        texts: dict[int, str] = {}
        todo = []  # (page_no, key, content)
        for page_no in page_nos:
            png = render_page_png(pdf_path, page_no, tcfg["dpi"])
            content = [{"type": "image", "source": {"type": "base64", "media_type": "image/png",
                                                    "data": base64.standard_b64encode(png).decode("ascii")}},
                       {"type": "text", "text": "이 페이지를 판독해 주세요."}]
            key = llm.cache_key(tcfg["model"], TRANSCRIBE_SYSTEM, content, TRANSCRIBE_SCHEMA, tcfg["effort"])
            cached = llm.cache_get(key)
            if cached is not None:
                texts[page_no] = cached["text"]
            else:
                todo.append((page_no, key, content))

        failed = []
        with ThreadPoolExecutor(max_workers=tcfg["workers"]) as pool:
            futures = {pool.submit(llm.request, tcfg["model"], TRANSCRIBE_SYSTEM, content, TRANSCRIBE_SCHEMA,
                                   tcfg["effort"]): (page_no, key) for page_no, key, content in todo}
            for n, future in enumerate(as_completed(futures), start=1):
                page_no, key = futures[future]
                try:  # 응답 해석과 캐시 기록은 DB를 쓰므로 주 스레드에서 한다
                    texts[page_no] = llm.parse_message(future.result(), key, "transcribe", tcfg["model"])["text"]
                except (LLMError, json.JSONDecodeError, KeyError) as e:
                    failed.append(page_no)
                    log.error("%s p.%d 판독 실패: %s", f["file_name"], page_no, str(e)[:200])
                    store.log("transcribe", "ERROR", f["file_name"], f"p.{page_no} 판독 실패: {str(e)[:200]}")
                if n % 20 == 0:
                    log.info("  … %d/%d쪽 (%s)", n, len(todo), llm.usage_summary())

        for page_no, text in texts.items():
            cleaned = clean_page_text(text, noise)
            n_chars = len(re.sub(r"\s+", "", cleaned))
            store.conn.execute("UPDATE pages SET text=?, n_chars=?, needs_ocr=0 WHERE file_hash=? AND page_no=?",
                               (cleaned, n_chars, f["file_hash"], page_no))
        if failed or len(texts) < f["n_pages"]:
            store.commit()
            log.warning("%s: %d/%d쪽만 판독됨 — 'transcribe'를 다시 실행하면 나머지만 이어서 판독합니다.",
                        f["file_name"], len(texts), f["n_pages"])
            continue
        store.conn.execute("UPDATE source_files SET extract_mode='transcribed' WHERE file_hash=?", (f["file_hash"],))
        store.commit()
        unreadable = sum(t.count("□") for t in texts.values())
        msg = f"{f['n_pages']}쪽 AI 판독 완료 (판독 불가 표시 □ {unreadable}개). 이후 발췌 검증은 'AI 판독문 대조'로 표시됨"
        log.info("%s: %s", f["file_name"], msg)
        store.log("transcribe", "WARNING", f["file_name"], msg)
    log.info(llm.usage_summary())
