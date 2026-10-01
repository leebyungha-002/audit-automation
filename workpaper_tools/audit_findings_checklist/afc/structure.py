"""지적사항 구조화(Claude API) + 환각 방지 검증(발췌·기준서 번호를 원문과 코드로 대조)."""
import base64
import difflib
import json
import re

from pydantic import ValidationError

from .db import Store, now
from .extract import render_page_png
from .llm import LLM, LLMError
from .models import StructuredFinding, finding_json_schema

SYSTEM_PROMPT = """당신은 금융감독원·한국공인회계사회가 공개한 회계 심사·감리 지적사례를 구조화된 레코드로 옮기는 보조자입니다.

결과물은 20년차 회계감사 실무자가 감사조서 작성의 참고자료로 사용하며, 모든 레코드는 사람이 원문과 대조해 검토합니다. 그리고 source_excerpt와 기준서 번호는 프로그램이 원문과 글자 단위로 대조합니다. 그래서 그럴듯한 추정보다 원문에 대한 충실성이 훨씬 중요합니다. 제공된 지적사례 원문에 근거한 내용만 작성하고, 원문에서 확인되지 않는 항목은 추측으로 채우지 말고 null(목록은 빈 목록)로 두세요.

입력은 지적사례 1건의 원문이며, [p.N]은 원본 문서의 페이지 번호입니다. 같은 페이지에 앞뒤 다른 사례의 일부가 섞여 있을 수 있는데, 이 경우 "대상 사례 제목"에 해당하는 사례만 다룹니다.

각 항목 작성 기준:
- is_finding: 입력이 실제 개별 지적사례이면 true. 목차·통계·개요처럼 지적사례가 아니면 false로 두고 나머지는 null 또는 빈 목록으로 둡니다.
- case_no: 원문에 적힌 공식 사례번호를 표기 그대로(예: "FSS/2512-10", "KICPA-2025-01"). 없으면 null.
- title: 원문의 사례 제목. 사례번호는 제외하고 제목만 적습니다.
- issue_area: 원문 머리말의 "쟁점 분야" 값. 없으면 null.
- finding_target: 지적 대상. 회사의 회계처리 위반만 지적되었으면 "회사", 감사인의 감사절차 소홀·독립성·법규 위반만 지적되었으면 "감사인", 둘 다 지적되었으면 "회사 및 감사인". 원문이 감사인에 대해 "지적하지 아니하였다"거나 감사보고서 감리를 실시하지 않았다고 밝힌 경우는 "회사"입니다.
- year: 원문의 "결정연도"/"결정일"에 적힌 연도. 없으면 null.
- fiscal_period: 원문의 "회계결산일" 값 그대로. 없으면 null.
- standards: 원문에 명시된 기준서만 적습니다. framework는 기업회계기준서(제1xxx호)·한국채택국제회계기준이면 "K-IFRS", 일반기업회계기준이면 "일반기업회계기준", 회계감사기준이면 "회계감사기준", 그 밖(외부감사법, 개념체계 등)은 "기타". number는 원문 표기에서 번호 부분만("제1115호", "제16장", "500"), name은 원문에 명칭이 있으면 그대로("고객과의 계약에서 생기는 수익"), 없으면 null. 문단 번호는 넣지 않고, 같은 기준서는 한 번만 적습니다.
- related_accounts: 이 지적과 직접 관련된 재무제표 계정과목. 원문에 나온 계정 명칭을 그대로 쓰고(예: 매출, 매출채권, 재고자산, 개발비, 대손충당금), 원문에 없는 계정을 연상해서 추가하지 않습니다. 계정과 무관한 사례(독립성 위반 등)는 빈 목록.
- finding_type: 허용된 값 중 이 사례의 핵심 위반 성격에 가장 가까운 것 하나.
- risk_summary: 어떤 상황에서 무엇을 어떻게 잘못 처리했고 재무제표에 어떤 영향이 있었는지를 원문 내용만으로 2~3문장으로 요약합니다.
- source_excerpt: 지적의 핵심을 보여주는 원문 구절을 1~3문장, 원문의 한 곳에서 연속된 그대로 옮깁니다. 요약하거나 단어를 바꾸거나 중간을 생략하지 않습니다. 원문 텍스트에 띄어쓰기가 빠져 있으면 자연스럽게 띄어 써도 되지만(대조 시 공백은 무시됩니다) 글자는 그대로 둡니다. 줄바꿈으로 끊긴 단어는 이어 붙입니다. 가급적 "회계기준 위반 지적 내용"(감사인 대상 사례는 감사절차 미흡 내용)에서 고릅니다.
- excerpt_page: source_excerpt가 시작되는 페이지 번호([p.N]의 N).
- audit_hint: 이 사례와 같은 위험에 대해 감사인이 확인해볼 만한 감사절차를 1~3문장으로 제안합니다. 이 항목만은 원문 인용이 아닌 당신의 제안이며 "AI 제안"으로 표시되어 검토됩니다. 다만 일반론이 아니라 이 사례의 사실관계와 원문의 시사점에 연결된 구체적인 절차여야 합니다. "(AI 제안)" 같은 머리말은 붙이지 않습니다.

risk_summary와 audit_hint는 감사조서 문체인 평서형 "~한다/~하였다"체로 통일합니다."""

EXCERPT_RETRY_NOTE = """
<재작성 요청>
직전에 작성한 source_excerpt가 원문과 글자 단위로 일치하지 않았습니다(일치율 {ratio}):
"{excerpt}"
원문의 한 곳에서 연속된 구절을 한 글자도 바꾸지 말고 다시 골라 옮겨 주세요. 길이를 줄여 한 문장만 옮겨도 됩니다. 나머지 항목은 그대로 유지합니다.
</재작성 요청>"""

VISION_NOTE = """
이 문서는 텍스트 추출 품질이 낮아 페이지 이미지를 함께 제공합니다. 추출 텍스트는 숫자·문장부호·띄어쓰기가 빠지거나 순서가 어긋나 있을 수 있으므로, 페이지 이미지를 정본으로 삼아 읽으세요. source_excerpt와 기준서 번호는 이미지에 보이는 그대로 옮깁니다."""

CASE_NO = re.compile(r"(FSS)\s*/\s*(\d{4})\s*-\s*(\d{2})|(KICPA)\s*-\s*(\d{4})\s*-\s*(\d{2})")
DECISION_YEAR = re.compile(r"결정\s*(?:연도|일)\s*:?\s*(\d{4})\s*년?")


def _ws(s: str) -> str:
    return re.sub(r"\s+", "", s)


def _letters(s: str) -> str:
    return re.sub(r"[^A-Za-z가-힣]", "", s)


def _hangul(s: str) -> str:
    return re.sub(r"[^가-힣]", "", s)


VERIFIED = ("일치", "일치(문자만)", "일치(한글만)")


def verify_excerpt(excerpt: str | None, pages: list[dict], fuzzy_threshold: float,
                   mode: str = "text") -> tuple[str, float, int | None]:
    """발췌가 원문에 실제로 있는지 확인. (상태, 일치율, 발췌가 있는 페이지)를 돌려준다.

    공백 무시 → 한글·영문만 → (vision 문서 한정) 한글만 순으로 대조한다.
    vision 문서는 숫자·영문·문장부호가 텍스트에서 빠져 있어 한글만 대조할 수 있다.
    """
    if not excerpt:
        return "발췌없음", 0.0, None
    full = "".join(p["text"] for p in pages)
    levels = [("일치", _ws), ("일치(문자만)", _letters)]
    if mode == "vision":
        levels.append(("일치(한글만)", _hangul))
    for label, norm in levels:
        target = norm(excerpt)
        if target and target in norm(full):
            head = target[:25]
            page = next((p["page"] for p in pages if head in norm(p["text"])), pages[0]["page"])
            return label, 1.0, page
    norm = levels[-1][1]
    target, source = norm(excerpt), norm(full)
    if not target:
        return "불일치", 0.0, None
    blocks = difflib.SequenceMatcher(None, target, source, autojunk=False).get_matching_blocks()
    ratio = sum(b.size for b in blocks) / len(target)
    return ("유사" if ratio >= fuzzy_threshold else "불일치"), round(ratio, 3), None


def verify_standards(standards: list[dict], pages: list[dict], mode: str) -> list[dict]:
    """기준서 번호의 숫자가 원문 텍스트에 있는지 확인한다."""
    full = _ws("".join(p["text"] for p in pages))
    checked = []
    for s in standards:
        digits = "".join(re.findall(r"\d+", s["number"]))
        if digits and digits in full:
            status = "원문확인"
        elif mode == "vision":
            status = "이미지판독(텍스트 대조 불가)"
        else:
            status = "원문 미확인"
        checked.append({**s, "check": status})
    return checked


def parse_case_no(text: str) -> str | None:
    m = CASE_NO.search(text)
    if not m:
        return None
    return f"FSS-{m.group(2)}-{m.group(3)}" if m.group(1) else f"KICPA-{m.group(5)}-{m.group(6)}"


def build_content(seg, pages: list[dict], src_file, cfg: dict) -> list[dict]:
    content: list[dict] = []
    if src_file["extract_mode"] == "vision":
        pdf_path = cfg["paths"]["input_pdfs"] / src_file["file_name"]
        for p in pages:
            png = render_page_png(pdf_path, p["page"], cfg["extract"]["vision_dpi"])
            content.append({"type": "text", "text": f"[p.{p['page']}] 페이지 이미지"})
            content.append({"type": "image", "source": {
                "type": "base64", "media_type": "image/png",
                "data": base64.standard_b64encode(png).decode("ascii")}})
    body = "\n\n".join(f"[p.{p['page']}]\n{p['text']}" for p in pages)
    content.append({"type": "text", "text": f"대상 사례 제목: {seg['title']}\n\n<원문>\n{body}\n</원문>"})
    return content


def run_structure(cfg: dict, store: Store, log, force: bool = False, name_filter: str | None = None,
                  limit: int | None = None) -> None:
    llm = LLM(cfg, store, log)
    schema = finding_json_schema(cfg["finding_types"])
    model = cfg["llm"]["structure_model"]
    files = {r["file_hash"]: r for r in store.query("SELECT * FROM source_files")}
    segments = store.query("SELECT * FROM segments ORDER BY file_hash, seq")
    done = {r["segment_id"] for r in store.query("SELECT segment_id FROM findings")}
    # 사람이 확정/제외한 레코드는 --force로도 덮어쓰지 않는다
    reviewed = {r["segment_id"] for r in store.query("SELECT segment_id FROM findings WHERE review_status<>'미검토'")}
    n_ok = n_skip = n_err = 0

    for seg in segments:
        src = files.get(seg["file_hash"])
        if src is None or (name_filter and name_filter not in src["file_name"]):
            continue
        if seg["segment_id"] in reviewed or (not force and seg["segment_id"] in done):
            continue
        if limit is not None and n_ok + n_skip + n_err >= limit:
            break

        pages = json.loads(seg["pages_json"])
        seg_text = seg["title"] + "\n" + "".join(p["text"] for p in pages)
        case_no = parse_case_no(seg_text)
        if case_no:  # 같은 사례가 다른 파일(연도별 사례집과 개별 파일)에 또 있으면 API 호출 전에 건너뛴다
            dup = store.query("SELECT source_file FROM findings WHERE finding_id=? AND segment_id<>?",
                              (case_no, seg["segment_id"]))
            if dup:
                msg = f"{case_no}: 이미 '{dup[0]['source_file']}'에서 등록된 사례 → 중복으로 건너뜀"
                log.warning(msg)
                store.log("structure", "WARNING", src["file_name"], msg)
                n_skip += 1
                continue
        system = SYSTEM_PROMPT + (VISION_NOTE if src["extract_mode"] == "vision" else "")
        try:
            content = build_content(seg, pages, src, cfg)
        except Exception as e:
            log.error("%s: 입력 구성 실패 (%s)", seg["segment_id"], e)
            store.log("structure", "ERROR", src["file_name"], f"{seg['segment_id']} 입력 구성 실패: {e}")
            n_err += 1
            continue

        finding, raw, error = None, None, None
        for attempt in range(cfg["llm"]["retries"] + 1):
            try:
                raw = llm.call_json("structure", model, system, content, schema, use_cache=(attempt == 0))
                finding = StructuredFinding.model_validate(raw)
                break
            except (json.JSONDecodeError, ValidationError) as e:
                error = f"출력 파싱/검증 실패({attempt + 1}차): {str(e)[:200]}"
                log.warning("%s: %s", seg["segment_id"], error)
            except LLMError as e:
                error = str(e)
                break
        if finding is None:
            log.error("%s (%s): 구조화 실패 — %s", seg["segment_id"], seg["title"], error)
            store.log("structure", "ERROR", src["file_name"], f"{seg['segment_id']} 구조화 실패: {error}")
            n_err += 1
            if error and ("인증" in error or "권한" in error or "크레딧" in error):
                log.error("API 사용 불가 상태로 보여 중단합니다.")
                break
            continue
        if not finding.is_finding:
            log.warning("%s (%s): 지적사례가 아닌 것으로 판단되어 제외", seg["segment_id"], seg["title"])
            store.log("structure", "WARNING", src["file_name"], f"{seg['segment_id']} 지적사례 아님으로 제외: {seg['title']}")
            n_skip += 1
            continue

        threshold, mode = cfg["verify"]["fuzzy_threshold"], src["extract_mode"]
        status, ratio, found_page = verify_excerpt(finding.source_excerpt, pages, threshold, mode)
        if status not in VERIFIED and finding.source_excerpt:
            # 발췌가 원문과 다르면 한 번 더 요청하고, 일치하는 쪽을 채택한다
            note = EXCERPT_RETRY_NOTE.format(ratio=ratio, excerpt=finding.source_excerpt)
            try:
                raw2 = llm.call_json("structure-retry", model, system, [*content, {"type": "text", "text": note}], schema)
                finding2 = StructuredFinding.model_validate(raw2)
                status2, ratio2, page2 = verify_excerpt(finding2.source_excerpt, pages, threshold, mode)
                if status2 in VERIFIED:
                    log.info("%s: 발췌 재작성으로 원문 일치 확보", seg["segment_id"])
                    finding, raw, status, ratio, found_page = finding2, raw2, status2, ratio2, page2
            except (json.JSONDecodeError, ValidationError, LLMError) as e:
                log.warning("%s: 발췌 재작성 실패 (%s)", seg["segment_id"], str(e)[:150])

        std_checked = verify_standards([s.model_dump() for s in finding.standards], pages, mode)
        year_in_text = DECISION_YEAR.search(seg_text)
        year = int(year_in_text.group(1)) if year_in_text else finding.year
        case_no = case_no or parse_case_no(finding.case_no or "")  # vision 문서는 AI가 이미지에서 읽은 번호 사용
        finding_id = case_no or f"F-{year or src['doc_year'] or 0}-{seg['segment_id']}"

        flags = []
        if status in ("유사", "불일치", "발췌없음"):
            flags.append(f"발췌 {status}({ratio})")
        flags += [f"기준서 {s['framework']} {s['number']} {s['check']}" for s in std_checked if s["check"] == "원문 미확인"]
        if year_in_text and finding.year and int(year_in_text.group(1)) != finding.year:
            flags.append(f"결정연도 불일치(원문 {year_in_text.group(1)} / AI {finding.year})")

        store.conn.execute("DELETE FROM findings WHERE segment_id=? AND finding_id<>?", (seg["segment_id"], finding_id))
        store.upsert("findings", {
            "finding_id": finding_id, "segment_id": seg["segment_id"], "case_no": case_no,
            "issue_area": finding.issue_area, "source_file": src["file_name"],
            "source_page": found_page or finding.excerpt_page or seg["page_start"],
            "source_page_end": seg["page_end"], "source_excerpt": finding.source_excerpt,
            "issuer": src["issuer"], "year": year, "fiscal_period": finding.fiscal_period,
            "title": finding.title or seg["title"], "finding_target": finding.finding_target,
            "standards": [s.model_dump() for s in finding.standards],
            "related_accounts": finding.related_accounts, "finding_type": finding.finding_type,
            "risk_summary": finding.risk_summary, "audit_hint": finding.audit_hint,
            "excerpt_status": status, "excerpt_match_ratio": ratio, "standards_check": std_checked,
            "segment_method": seg["method"], "extract_mode": src["extract_mode"],
            "llm_model": model, "prompt_version": cfg["llm"]["prompt_version"],
            "llm_original_json": raw, "review_status": "미검토", "created_at": now(),
        })
        store.commit()
        n_ok += 1
        if flags:
            msg = f"{finding_id} 검증 플래그: " + "; ".join(flags)
            log.warning(msg)
            store.log("structure", "WARNING", src["file_name"], msg)
        else:
            log.info("%s 구조화 완료 (%s)", finding_id, finding.title)

    log.info("구조화 결과: 성공 %d건, 제외 %d건, 오류 %d건", n_ok, n_skip, n_err)
    log.info(llm.usage_summary())
