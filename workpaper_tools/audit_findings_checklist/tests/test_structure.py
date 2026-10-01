"""구조화와 환각 방지 검증 테스트 (API는 가짜 LLM으로 대체)."""
import json

import pytest

import afc.structure as structure
from afc.structure import parse_case_no, run_structure, verify_excerpt, verify_standards
from conftest import CASE_TEXT, GOOD_EXCERPT, FakeLLM, add_finding, add_segment, add_source, llm_output

PAGES = [{"page": 3, "text": "1. 회사의 회계처리\n내용"}, {"page": 4, "text": CASE_TEXT}]


def use_fake(monkeypatch, responses):
    fake = FakeLLM(responses)
    monkeypatch.setattr(structure, "LLM", lambda cfg, store, log: fake)
    return fake


@pytest.mark.parametrize("excerpt, mode, status", [
    (GOOD_EXCERPT, "text", "일치"),
    (GOOD_EXCERPT.replace(" ", ""), "text", "일치"),                       # 띄어쓰기 차이는 무시
    (GOOD_EXCERPT.replace("매출로", "매출(로)"), "text", "일치(문자만)"),   # 문장부호 차이
    ("회사는 계약금액을 허위로 증액하여 이를 자기자본으로 인식하고 매출 및 영업이익을 과대계상하였다.", "text", "유사"),
    ("원문에 전혀 없는 문장을 지어냈다.", "text", "불일치"),
    (None, "text", "발췌없음"),
])
def test_verify_excerpt_levels(excerpt, mode, status):
    assert verify_excerpt(excerpt, PAGES, 0.85, mode)[0] == status


def test_verify_excerpt_reports_page_where_excerpt_is():
    assert verify_excerpt(GOOD_EXCERPT, PAGES, 0.85)[2] == 4


def test_vision_documents_compare_hangul_only():
    pages = [{"page": 1, "text": "회사가 년 사에 공급한 상품은 통제가 이전되었다고 보기 어려웠다"}]  # 숫자·영문이 빠진 추출
    excerpt = "회사가 ’x1년 A사에 공급한 상품은 통제가 이전되었다고 보기 어려웠다"
    assert verify_excerpt(excerpt, pages, 0.85, "text")[0] not in structure.VERIFIED
    assert verify_excerpt(excerpt, pages, 0.85, "vision")[0] == "일치(한글만)"


def test_transcribed_documents_are_marked_as_checked_against_ai_transcript():
    status = verify_excerpt(GOOD_EXCERPT, PAGES, 0.85, "transcribed")[0]
    assert status == "일치(AI 판독문)" and status in structure.VERIFIED
    assert verify_excerpt("원문에 전혀 없는 문장을 지어냈다.", PAGES, 0.85, "transcribed")[0] == "불일치"


def test_verify_standards_flags_numbers_missing_from_source():
    standards = [{"framework": "K-IFRS", "number": "제1018호", "name": None},
                 {"framework": "K-IFRS", "number": "제9999호", "name": None}]
    assert [s["check"] for s in verify_standards(standards, PAGES, "text")] == ["원문확인", "원문 미확인"]
    assert verify_standards(standards, PAGES, "vision")[1]["check"] == "이미지판독(텍스트 대조 불가)"


def test_parse_case_no():
    assert parse_case_no("감리지적사례 FSS/2106-03 : 매출") == "FSS-2106-03"
    assert parse_case_no("심사·감리지적사례 KICPA-2025-01") == "KICPA-2025-01"
    assert parse_case_no("사례번호 없음") is None


def test_structure_stores_finding_with_official_id_and_verification(env, monkeypatch):
    cfg, store, log = env
    add_segment(store, add_source(store))
    use_fake(monkeypatch, [llm_output()])
    run_structure(cfg, store, log)

    f = store.query("SELECT * FROM findings")[0]
    assert f["finding_id"] == "FSS-2106-03"
    assert f["year"] == 2020 and f["excerpt_status"] == "일치" and f["review_status"] == "미검토"
    assert json.loads(f["standards_check"])[0]["check"] == "원문확인"
    assert json.loads(f["llm_original_json"])["source_excerpt"] == GOOD_EXCERPT


def test_misquoted_excerpt_is_retried_and_corrected(env, monkeypatch):
    cfg, store, log = env
    add_segment(store, add_source(store))
    bad = llm_output(source_excerpt=GOOD_EXCERPT.replace("매출로", "자기자본으로"))
    fake = use_fake(monkeypatch, [bad, llm_output()])
    run_structure(cfg, store, log)

    assert [c["purpose"] for c in fake.calls] == ["structure", "structure-retry"]
    assert store.query("SELECT excerpt_status FROM findings")[0]["excerpt_status"] == "일치"


def test_excerpt_that_stays_wrong_is_stored_with_flag(env, monkeypatch):
    cfg, store, log = env
    add_segment(store, add_source(store))
    bad = llm_output(source_excerpt="원문에 전혀 없는 문장을 지어냈다.")
    use_fake(monkeypatch, [bad, bad])
    run_structure(cfg, store, log)

    assert store.query("SELECT excerpt_status FROM findings")[0]["excerpt_status"] == "불일치"
    assert store.query("SELECT 1 FROM process_log WHERE level='WARNING' AND message LIKE '%검증 플래그%'")


def test_already_structured_segments_are_not_sent_again(env, monkeypatch):
    cfg, store, log = env
    add_segment(store, add_source(store))
    fake = use_fake(monkeypatch, [llm_output()])
    run_structure(cfg, store, log)
    run_structure(cfg, store, log)  # 증분: 새 구간이 없으므로 호출 없음
    assert len(fake.calls) == 1


def test_same_case_in_another_file_is_skipped_without_api_call(env, monkeypatch):
    cfg, store, log = env
    add_segment(store, add_source(store, "개별.hwp", "a" * 64))
    add_segment(store, add_source(store, "사례집.pdf", "b" * 64))
    fake = use_fake(monkeypatch, [llm_output()])
    run_structure(cfg, store, log)

    assert len(fake.calls) == 1 and len(store.query("SELECT 1 FROM findings")) == 1
    assert "중복" in store.query("SELECT reason FROM segment_skips")[0]["reason"]


def test_non_finding_segment_is_recorded_and_not_retried(env, monkeypatch):
    cfg, store, log = env
    add_segment(store, add_source(store), title="통계", text="Ⅰ. 회계심사 실시 통계\n건수 요약")
    fake = use_fake(monkeypatch, [llm_output(is_finding=False)])
    run_structure(cfg, store, log)
    run_structure(cfg, store, log)

    assert len(fake.calls) == 1 and not store.query("SELECT 1 FROM findings")
    assert structure.estimate(cfg, store, log)[0] == 0


def test_reviewed_findings_survive_forced_rerun(env, monkeypatch):
    cfg, store, log = env
    segment_id = add_segment(store, add_source(store))
    add_finding(store, "FSS-2106-03", segment_id, review_status="확정", risk_summary="감사인이 고친 요약")
    fake = use_fake(monkeypatch, [llm_output()])
    run_structure(cfg, store, log, force=True)

    assert not fake.calls
    assert store.query("SELECT risk_summary FROM findings")[0]["risk_summary"] == "감사인이 고친 요약"


def test_batch_submits_pending_segments_and_stores_results(env, monkeypatch):
    cfg, store, log = env
    add_segment(store, add_source(store))
    fake = use_fake(monkeypatch, [llm_output()])
    run_structure(cfg, store, log, batch=True, wait_minutes=0)

    assert fake.submitted == [("batch-1", ["aaaaaaaa-001"])]
    assert store.query("SELECT finding_id FROM findings")[0]["finding_id"] == "FSS-2106-03"
    assert store.query("SELECT status FROM batches")[0]["status"] == "processed"


def test_segments_in_an_unfinished_batch_are_not_submitted_again(env, monkeypatch):
    cfg, store, log = env
    add_segment(store, add_source(store))
    fake = use_fake(monkeypatch, [llm_output()])
    monkeypatch.setattr(fake, "batch_ended", lambda batch_id: (False, "처리 중 1"))
    run_structure(cfg, store, log, batch=True, wait_minutes=0)   # 제출만 되고 끝나지 않음
    run_structure(cfg, store, log, batch=True, wait_minutes=0)   # 다시 실행해도 재제출하지 않음
    assert len(fake.submitted) == 1 and structure.estimate(cfg, store, log)[0] == 0

    monkeypatch.setattr(fake, "batch_ended", lambda batch_id: (True, "완료"))
    run_structure(cfg, store, log, batch=True, wait_minutes=0)   # 이어받기
    assert len(store.query("SELECT 1 FROM findings")) == 1


def test_invalid_llm_output_is_retried_then_logged_as_error(env, monkeypatch):
    cfg, store, log = env
    add_segment(store, add_source(store))
    broken = llm_output(finding_target="잘못된 값")
    fake = use_fake(monkeypatch, [broken] * (cfg["llm"]["retries"] + 1))
    run_structure(cfg, store, log)

    assert len(fake.calls) == cfg["llm"]["retries"] + 1
    assert not store.query("SELECT 1 FROM findings")
    assert store.query("SELECT 1 FROM process_log WHERE level='ERROR' AND message LIKE '%구조화 실패%'")
