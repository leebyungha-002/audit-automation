"""추출(증분 처리 포함)과 분할 규칙 테스트."""
import re

from afc.extract import clean_page_text, decide_mode, run_extract
from afc.segment import _pick_profile, _split_anchor, _split_page_title, run_segment
from conftest import CASE_TEXT, make_pdf


def test_extract_keeps_page_numbers_and_skips_processed_files(env):
    cfg, store, log = env
    make_pdf(cfg["paths"]["input_pdfs"] / "FSS2106_03.pdf", [CASE_TEXT, "5. 시사점\n한계기업의 매출 급증에 주의한다."])
    run_extract(cfg, store, log)

    pages = store.query("SELECT page_no, text FROM pages ORDER BY page_no")
    assert [p["page_no"] for p in pages] == [1, 2]
    assert "FSS/2106-03" in pages[0]["text"] and "시사점" in pages[1]["text"]
    first_run = store.query("SELECT processed_at, issuer FROM source_files")[0]
    assert first_run["issuer"] == "금감원"

    store.conn.execute("UPDATE source_files SET processed_at='이전 실행'")
    run_extract(cfg, store, log)  # 해시가 같으므로 다시 처리하지 않는다
    assert store.query("SELECT processed_at FROM source_files")[0]["processed_at"] == "이전 실행"


def test_excluded_and_scanned_files(env):
    cfg, store, log = env
    make_pdf(cfg["paths"]["input_pdfs"] / "2024년_강의_교안.pdf", [CASE_TEXT])
    make_pdf(cfg["paths"]["input_pdfs"] / "scan.pdf", ["", CASE_TEXT])
    run_extract(cfg, store, log)

    names = [r["file_name"] for r in store.query("SELECT file_name FROM source_files")]
    assert names == ["scan.pdf"]  # '교안'은 설정에서 제외
    assert [r["needs_ocr"] for r in store.query("SELECT needs_ocr FROM pages ORDER BY page_no")] == [1, 0]
    assert store.query("SELECT 1 FROM process_log WHERE level='WARNING' AND message LIKE '%OCR%'")


def test_changed_file_replaces_old_version_but_keeps_reviewed_findings(env):
    from conftest import add_finding

    cfg, store, log = env
    pdf = cfg["paths"]["input_pdfs"] / "사례집.pdf"
    make_pdf(pdf, [CASE_TEXT])
    run_extract(cfg, store, log)
    run_segment(cfg, store, log)
    old_segment = store.query("SELECT segment_id FROM segments")[0]["segment_id"]
    add_finding(store, "FSS-2106-03", old_segment, review_status="확정")

    make_pdf(pdf, [CASE_TEXT.replace("2106-03", "2106-04")])  # 같은 이름, 다른 내용
    run_extract(cfg, store, log)

    assert len(store.query("SELECT 1 FROM source_files")) == 1
    assert "2106-04" in store.query("SELECT text FROM pages")[0]["text"]
    assert store.query("SELECT review_status FROM findings")[0]["review_status"] == "확정"


def test_clean_page_text_removes_headers_and_private_use_chars(env):
    cfg, _, _ = env
    noise = [re.compile(p) for p in cfg["extract"]["noise_lines"]]
    text = "- 12 -\nSlide 3\n쟁점분야: 매출\n\n본문"
    assert clean_page_text(text, noise) == "쟁점분야: 매출\n본문"


def test_garbled_text_layer_is_detected():
    from afc.extract import readable_ratio
    assert readable_ratio("회사는 매출 15억원을 과대계상하였다(K-IFRS 제1115호).") > 0.9
    assert readable_ratio("ᵘᶬエㄬ㋤ᛈㅸ⏌ᬬ゘⎔ジ䀰㶀ㅜ㎨⍤㮝㽜㋤ぼⳔ⪔") < 0.2


def test_decide_mode(env):
    ecfg = env[0]["extract"]
    assert decide_mode("a.pdf", images_per_page=15.0, ecfg=ecfg) == "vision"
    assert decide_mode("a.pdf", images_per_page=0.1, ecfg=ecfg) == "text"
    assert decide_mode("a.pdf", 0.0, {**ecfg, "force_vision": ["a.pdf"]}) == "vision"


def _profile(cfg, name):
    return next(p for p in cfg["segment_profiles"] if p["name"] == name)


def test_single_fss_case_is_one_segment(env):
    cfg = env[0]
    pages = [{"page": 1, "text": CASE_TEXT}, {"page": 2, "text": "5. 시사점\n주의한다."}]
    prof = _pick_profile(pages, cfg["segment_profiles"])
    assert prof["name"] == "fss_single_case"
    segments = _split_anchor(pages, prof)
    assert len(segments) == 1
    assert segments[0]["title"] == "감리지적사례 FSS/2106-03 : 매출 허위계상"
    assert [p["page"] for p in segments[0]["pages"]] == [1, 2]


def test_kicpa_casebook_splits_cases_that_share_a_page(env):
    cfg = env[0]
    pages = [
        {"page": 1, "text": "목차\n㉮ KICPA-2025-01 공사수익 오류"},
        {"page": 7, "text": "심사·감리지적사례 KICPA-2025-01\n: 공사수익과 공사원가 오류\n1. 회사의 회계처리\n내용A"},
        {"page": 8, "text": "5. 시사점\n끝A\n심사·감리지적사례 KICPA-2025-02\n: 분양수익 오류\n1. 회사의 회계처리\n내용B"},
    ]
    prof = _pick_profile(pages, cfg["segment_profiles"])
    assert prof["name"] == "kicpa_case"
    first, second = _split_anchor(pages, prof)
    assert first["title"] == "심사·감리지적사례 KICPA-2025-01 : 공사수익과 공사원가 오류"
    assert [p["page"] for p in first["pages"]] == [7, 8]
    assert "끝A" in first["pages"][1]["text"] and "내용B" not in first["pages"][1]["text"]
    assert [p["page"] for p in second["pages"]] == [8]


def test_fss_report_title_comes_from_lines_above_anchor(env):
    cfg = env[0]
    pages = [{"page": 6, "text": "Ⅱ. 주요 지적사례\n1\n금융상품 계정분류 오류\n1. 회사의 회계처리\n내용"},
             {"page": 8, "text": "2\n미지급장려금 과소계상\n1. 회사의 회계처리\n내용"},
             {"page": 10, "text": "3\n재고자산 과대계상\n1. 회사의 회계처리\n내용"}]
    segments = _split_anchor(pages, _profile(cfg, "fss_report_case"))
    assert [s["title"] for s in segments] == ["금융상품 계정분류 오류", "미지급장려금 과소계상", "재고자산 과대계상"]


def test_slides_group_consecutive_pages_and_skip_contents_page(env):
    cfg = env[0]
    pages = [
        {"page": 2, "text": "감리사례1: A사[매출허위계상]\n감리사례2: C사[재고과대]"},  # 목차 (키 2개)
        {"page": 9, "text": "감리사례1: A사[매출허위계상]\n1. 사실관계"},
        {"page": 10, "text": "감리사례1: A사[매출허위계상]\n2. 지적내용"},
        {"page": 16, "text": "감리사례2: C사[재고과대]\n1. 사실관계"},
    ]
    segments = _split_page_title(pages, _profile(cfg, "slide_case"))
    assert [[p["page"] for p in s["pages"]] for s in segments] == [[9, 10], [16]]


def test_no_profile_for_unstructured_text(env):
    assert _pick_profile([{"page": 1, "text": "회계심사 통계 자료입니다."}], env[0]["segment_profiles"]) is None
