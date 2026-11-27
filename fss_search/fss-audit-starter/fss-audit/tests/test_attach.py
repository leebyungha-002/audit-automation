import pytest

import attach

ITEM = {"contentId": "1", "atchfileNm": "../../사례:1.hwp",
        "atchfileUrl": "https://www.fss.or.kr/file?a=1&amp;b=2"}


class Fake:
    def __init__(self, data=b"", status=200, location=None):
        self.data, self.status_code = data, status
        self.headers = {"Location": location} if location else {}

    def raise_for_status(self):
        pass

    def iter_content(self, size):
        yield self.data


def test_save_and_keep_versions(tmp_path, monkeypatch):
    monkeypatch.setattr(attach.requests, "get", lambda *a, **k: Fake(b"v1"))
    first = attach.download("주석공시", ITEM, str(tmp_path))
    assert first.read_bytes() == b"v1" and tmp_path in first.resolve().parents
    assert attach.download("주석공시", ITEM, str(tmp_path)) == first
    monkeypatch.setattr(attach.requests, "get", lambda *a, **k: Fake(b"v2"))
    assert attach.download("주석공시", ITEM, str(tmp_path)) != first
    assert first.read_bytes() == b"v1"


@pytest.mark.parametrize("url,response", [
    ("https://evil.example/x", Fake(b"x")),
    ("https://www.fss.or.kr/x", Fake(status=302, location="https://evil.example/x")),
    ("https://www.fss.or.kr/x", Fake(b"<!DOCTYPE html><p>error")),
    ("https://www.fss.or.kr/x", Fake(b"")),
])
def test_rejects(tmp_path, monkeypatch, url, response):
    monkeypatch.setattr(attach.requests, "get", lambda *a, **k: response)
    with pytest.raises(attach.DownloadError):
        attach.download("a", {**ITEM, "atchfileUrl": url}, str(tmp_path))
