from streamlit.testing.v1 import AppTest

import fss_client
import store
from test_core import ITEM


def test_search_screen(tmp_path, monkeypatch):
    db = str(tmp_path / "t.db")
    con = store.connect(db)
    store.save(con, "주석공시", [ITEM])
    con.close()
    monkeypatch.setenv("FSS_DB_PATH", db)
    at = AppTest.from_file("../app.py", default_timeout=30).run()
    at.text_input[0].input("1118").run()
    assert not at.exception
    assert len(at.dataframe[0].value) == 1


def test_collect_button(tmp_path, monkeypatch):
    monkeypatch.setenv("FSS_DB_PATH", str(tmp_path / "t.db"))
    monkeypatch.setenv("FSS_API_KEY", "dummy")
    monkeypatch.setattr(fss_client, "fetch", lambda *a: [ITEM])
    at = AppTest.from_file("../app.py", default_timeout=30).run()
    at.number_input[0].set_value(at.number_input[1].value).run()
    at.multiselect[0].set_value(["주석공시"]).run()
    at.button[0].click().run()
    assert not at.exception
    assert at.success[0].value == "수집 완료"
    assert "1건" in at.caption[0].value


def test_attachment_button(tmp_path, monkeypatch):
    import attach
    db = str(tmp_path / "t.db")
    con = store.connect(db)
    store.save(con, "주석공시", [{**ITEM, "atchfileUrl": "https://www.fss.or.kr/f",
                                 "atchfileNm": "a.hwp"}])
    con.close()
    saved = tmp_path / "a.hwp"
    saved.write_bytes(b"data")
    monkeypatch.setenv("FSS_DB_PATH", db)
    monkeypatch.setattr(attach, "download", lambda *a: saved)
    at = AppTest.from_file("../app.py", default_timeout=30).run()
    at.text_input[0].input("1118").run()
    next(b for b in at.button if b.label == "첨부파일 가져오기").click().run()
    assert not at.exception
    assert str(saved) in at.success[0].value


def test_collect_continues_after_api_error(tmp_path, monkeypatch):
    monkeypatch.setenv("FSS_DB_PATH", str(tmp_path / "t.db"))
    monkeypatch.setenv("FSS_API_KEY", "dummy")

    def fetch(api, *a):
        if api == "주석공시":
            raise fss_client.FssApiError("호출 실패: HTTPError")
        return [ITEM]
    monkeypatch.setattr(fss_client, "fetch", fetch)
    at = AppTest.from_file("../app.py", default_timeout=30).run()
    at.number_input[0].set_value(at.number_input[1].value).run()
    at.multiselect[0].set_value(["주석공시", "외부감사 FAQ"]).run()
    at.button[0].click().run()
    assert not at.exception
    assert "주석공시" in at.warning[0].value
    assert "1건" in at.caption[0].value  # 외부감사 FAQ는 저장됨
