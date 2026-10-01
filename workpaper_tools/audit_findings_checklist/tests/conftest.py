"""테스트 공통 준비물. 실제 DB·원문·API를 건드리지 않도록 임시 폴더와 가짜 LLM을 쓴다."""
import logging
import sys
from pathlib import Path

import pymupdf
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from afc.config import load_config  # noqa: E402
from afc.db import Store, now  # noqa: E402

CASE_TEXT = (
    "감리지적사례 FSS/2106-03 : 매출 허위계상\n"
    "▣ 쟁점 분야: 매출\n"
    "▣ 관련 기준: 기업회계기준서 제1018호\n"
    "▣ 결정일: 2020년\n"
    "1. 회사의 회계처리\n"
    "A사는 종속회사와의 용역계약 금액을 임의로 증액하였다.\n"
    "2. 회계기준 위반 지적 내용\n"
    "회사는 계약금액을 허위로 증액하여 이를 매출로 인식하고 매출 및 영업이익을 과대계상하였다.\n"
)
GOOD_EXCERPT = "회사는 계약금액을 허위로 증액하여 이를 매출로 인식하고 매출 및 영업이익을 과대계상하였다."


@pytest.fixture
def env(tmp_path):
    """(cfg, store, log). 경로는 모두 임시 폴더를 가리킨다."""
    cfg = load_config()
    cfg["paths"] = {"input_pdfs": tmp_path / "input", "db": tmp_path / "findings.db",
                    "output": tmp_path / "output", "logs": tmp_path / "logs"}
    for key in ("input_pdfs", "output", "logs"):
        cfg["paths"][key].mkdir()
    store = Store(cfg["paths"]["db"])
    yield cfg, store, logging.getLogger("afc-test")
    store.close()


def make_pdf(path: Path, pages: list[str]) -> None:
    doc = pymupdf.open()
    for text in pages:
        page = doc.new_page()
        for i, line in enumerate(text.splitlines()):
            page.insert_text((50, 60 + 20 * i), line, fontname="korea", fontsize=10)
    doc.save(path)
    doc.close()


def add_source(store: Store, name: str = "FSS2106_03.pdf", file_hash: str = "aaaaaaaa" * 8,
               mode: str = "text", issuer: str = "금감원") -> str:
    store.upsert("source_files", {"file_hash": file_hash, "file_name": name, "issuer": issuer, "doc_year": 2021,
                                  "n_pages": 1, "extract_mode": mode, "processed_at": now()})
    store.commit()
    return file_hash


def add_segment(store: Store, file_hash: str, seq: int = 1, title: str = "감리지적사례 FSS/2106-03 : 매출 허위계상",
                text: str = CASE_TEXT) -> str:
    segment_id = f"{file_hash[:8]}-{seq:03d}"
    store.upsert("segments", {"segment_id": segment_id, "file_hash": file_hash, "seq": seq, "title": title,
                              "page_start": 1, "page_end": 1, "pages_json": [{"page": 1, "text": text}],
                              "method": "규칙", "profile": "fss_single_case"})
    store.commit()
    return segment_id


def llm_output(**overrides) -> dict:
    """구조화 LLM이 돌려줄 법한 정상 출력."""
    data = {
        "is_finding": True, "case_no": "FSS/2106-03", "title": "매출 허위계상", "issue_area": "매출",
        "finding_target": "회사", "year": 2020, "fiscal_period": None,
        "standards": [{"framework": "K-IFRS", "number": "제1018호", "name": None}],
        "related_accounts": ["매출", "매출채권"], "finding_type": "허위·가공 계상",
        "risk_summary": "계약금액을 허위로 증액하여 매출을 과대계상하였다.",
        "source_excerpt": GOOD_EXCERPT, "excerpt_page": 1,
        "audit_hint": "계약금액 변경 근거를 확인한다.",
    }
    data.update(overrides)
    return data


class FakeLLM:
    """API를 부르지 않는 LLM 대역. responses를 순서대로 돌려주고 받은 입력을 기록한다."""

    def __init__(self, responses: list[dict]):
        self.responses = list(responses)
        self.calls: list[dict] = []
        self.cfg = load_config()["llm"]

    def call_json(self, purpose, model, system, content, schema, use_cache=True):
        self.calls.append({"purpose": purpose, "model": model, "system": system, "content": content})
        return self.responses.pop(0)

    def request(self, model, system, content, schema, effort=None):
        """즉시 처리 경로의 첫 요청. 응답 '메시지' 자리에 준비된 dict를 그대로 돌려준다."""
        return self.call_json("structure", model, system, content, schema)

    def usage_summary(self) -> str:
        return f"가짜 호출 {len(self.calls)}회"

    # ── 배치 대역: 제출한 요청마다 responses에서 하나씩 꺼내 결과로 돌려준다 ──
    def cache_key(self, model, system, content, schema) -> str:
        return "key"

    def cache_get(self, key):
        return None

    def submit_batch(self, model, schema, items) -> str:
        self.submitted = getattr(self, "submitted", [])
        batch_id = f"batch-{len(self.submitted) + 1}"
        self.submitted.append((batch_id, [cid for cid, _, _ in items]))
        return batch_id

    def batch_ended(self, batch_id):
        return True, "완료"

    def batch_results(self, batch_id):
        from types import SimpleNamespace
        for cid in dict(self.submitted)[batch_id]:
            yield SimpleNamespace(custom_id=cid, result=SimpleNamespace(type="succeeded", message=self.responses.pop(0)))

    def parse_message(self, message, key, purpose, requested_model, discount=1.0):
        return message


def add_finding(store: Store, finding_id: str, segment_id: str, **overrides) -> None:
    row = {
        "finding_id": finding_id, "segment_id": segment_id, "case_no": finding_id, "issue_area": "매출",
        "source_file": "FSS2106_03.pdf", "source_page": 1, "source_page_end": 1, "source_excerpt": GOOD_EXCERPT,
        "issuer": "금감원", "year": 2020, "fiscal_period": None, "title": "매출 허위계상", "finding_target": "회사",
        "standards": [{"framework": "K-IFRS", "number": "제1018호", "name": None}],
        "related_accounts": ["매출", "매출채권"], "finding_type": "허위·가공 계상",
        "risk_summary": "계약금액을 허위로 증액하여 매출을 과대계상하였다.", "audit_hint": "계약금액 변경 근거를 확인한다.",
        "excerpt_status": "일치", "excerpt_match_ratio": 1.0,
        "standards_check": [{"framework": "K-IFRS", "number": "제1018호", "name": None, "check": "원문확인"}],
        "segment_method": "규칙", "extract_mode": "text", "llm_model": "test", "prompt_version": "t",
        "llm_original_json": {"original": True}, "review_status": "미검토", "created_at": now(),
    }
    row.update(overrides)
    store.upsert("findings", row)
    store.commit()
