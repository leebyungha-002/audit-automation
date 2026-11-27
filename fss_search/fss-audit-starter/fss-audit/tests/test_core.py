import pytest

import fss_client
import store

ITEM = {"contentId": "217244", "subject": "K-IFRS 제1118호 영향 관련 사전 주석공시 모범사례",
        "contentKor": "&lt;p&gt;&lt;span&gt;재무제표 표시와 공시&amp;nbsp;안내&lt;/span&gt;&lt;/p&gt;",
        "regDate": "2026-04-15 09:33:51", "originUrl": "https://example.invalid"}
OK = {"reponse": {"resultCode": "1", "resultCnt": 1, "result": [ITEM]}}


def test_parse_and_clean():
    assert fss_client.parse_response(OK) == [ITEM]
    assert fss_client.clean_html(ITEM["contentKor"]) == "재무제표 표시와 공시 안내"


def test_count_mismatch_raises():
    bad = {"reponse": {**OK["reponse"], "resultCnt": 2}}
    with pytest.raises(fss_client.FssApiError):
        fss_client.parse_response(bad)


def test_versioning_and_search():
    con = store.connect(":memory:")
    assert store.save(con, "주석공시", [ITEM]) == 1
    assert store.save(con, "주석공시", [ITEM]) == 0
    assert store.save(con, "주석공시", [{**ITEM, "subject": "정정"}]) == 1
    assert len(store.search(con, "재무제표")) == 2


def test_key_not_leaked(monkeypatch):
    def boom(*a, **k):
        raise fss_client.requests.ConnectionError("url?authKey=SECRET")
    monkeypatch.setattr(fss_client.requests, "get", boom)
    with pytest.raises(fss_client.FssApiError) as e:
        fss_client.fetch("주석공시", "2026-01-01", "2026-12-31", "SECRET")
    assert "SECRET" not in str(e.value)


def test_search_extra_fields_and_filter():
    con = store.connect(":memory:")
    store.save(con, "주석공시", [ITEM])
    case = {k: v for k, v in ITEM.items() if k != "contentKor"}  # 홈페이지 표기는 contentsKor
    store.save(con, "심사·감리지적사례", [{**case, "contentId": "9", "subject": "재고 과대계상",
                                     "contentsKor": "평가손실 미인식", "relStand": "K-IFRS 1002"}])
    assert len(store.search(con, "1002")) == 1          # 관련기준서로 검색
    assert len(store.search(con, "평가손실")) == 1       # contentsKor 표기도 본문으로 저장
    assert len(store.search(con, "1118", "심사·감리지적사례")) == 0
    assert len(store.search(con, "1118", "주석공시")) == 1


def test_snippet():
    body = "가" * 100 + "수익인식" + "나" * 100
    s = store.snippet(body, "수익인식", 10)
    assert s == "…" + "가" * 10 + "수익인식" + "나" * 10 + "…"
    assert store.snippet("짧은 본문", "없음") == "짧은 본문"


def test_collect_skips_failed_api_only(monkeypatch):
    calls = []

    def fetch(api, start, end, key):
        calls.append((api, start[:4]))
        if api == "A":
            raise fss_client.FssApiError("호출 실패: HTTPError")
        return [ITEM]
    monkeypatch.setattr(fss_client, "fetch", fetch)
    con = store.connect(":memory:")
    out = list(store.collect(con, "k", 2024, 2025, ["A", "B"], pause=0))
    assert [(a, y, e is None) for a, y, *_, e in out] == [("A", 2024, False), ("B", 2024, True),
                                                           ("B", 2025, True)]
    assert ("A", "2025") not in calls
