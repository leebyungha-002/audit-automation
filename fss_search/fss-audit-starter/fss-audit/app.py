"""실행: streamlit run app.py"""
import os
from datetime import date
from pathlib import Path

import streamlit as st
from dotenv import load_dotenv

import attach
import fss_client
import store

load_dotenv()
st.set_page_config(page_title="FSS 감사 참고자료 검색", layout="wide")
st.title("FSS 감사 참고자료 검색")

con = store.connect(os.getenv("FSS_DB_PATH", "fss.db"))
APIS = list(fss_client.ENDPOINTS)

# 수집·저장
with st.sidebar:
    st.header("자료 수집")
    this_year = date.today().year
    first = int(st.number_input("시작 연도", 2000, this_year, 2018))
    last = int(st.number_input("종료 연도", 2000, this_year, this_year))
    targets = st.multiselect("수집할 자료", APIS, default=APIS)
    if st.button("수집 및 저장"):
        key = os.getenv("FSS_API_KEY")
        if not key:
            st.error(".env 파일에 FSS_API_KEY를 설정하세요.")
        elif first > last:
            st.error("시작 연도가 종료 연도보다 늦습니다.")
        elif not targets:
            st.error("수집할 자료를 하나 이상 고르세요.")
        else:
            failed = []
            with st.spinner("수집 중"):
                for api, year, got, changed, err in store.collect(con, key, first, last, targets):
                    if err:
                        failed.append(api)
                        st.error(f"{api} {year}: {err} → 이 자료는 건너뜀")
                    else:
                        st.write(f"{api} {year}: 수신 {got}건, 신규/변경 {changed}건")
            if failed:
                st.warning(f"일부 실패: {', '.join(failed)}")
            else:
                st.success("수집 완료")

# 저장 현황
counts = dict(con.execute(
    "SELECT api, COUNT(DISTINCT content_id) FROM docs GROUP BY api").fetchall())
total = con.execute("SELECT COUNT(*) FROM docs").fetchone()[0]
st.caption(f"저장된 자료(이전 버전 포함): {total}건 · 검색 대상: 제목·본문·관련기준서·쟁점분야 "
           f"(첨부파일 제외)")
if counts:
    st.caption(" · ".join(f"{a} {n}건" for a, n in counts.items()))

# 키워드 검색
col1, col2 = st.columns([3, 1])
keyword = col1.text_input("검색어", placeholder="예: 1118, 수익인식, 재고자산").strip()
scope = col2.selectbox("구분", ["전체"] + APIS)
api_filter = None if scope == "전체" else scope

if keyword:
    rows = [
        {"구분": a, "등록일": (reg or "")[:10], "제목": subj,
         "내용": store.snippet(body, keyword), "원문": url, "수집일시": fetched}
        for a, reg, subj, url, fetched, body in store.search(con, keyword, api_filter)
    ]
    st.write(f"검색 결과 {len(rows)}건")
    if rows:
        st.dataframe(
            rows,
            hide_index=True,
            column_config={"원문": st.column_config.LinkColumn("원문", display_text="열기")},
        )

    # 첨부파일 다운로드
    files = store.search_attachments(con, keyword, api_filter)
    if files:
        st.subheader("첨부파일")
        index = st.selectbox(
            "받을 파일",
            range(len(files)),
            format_func=lambda i: f"[{files[i][0]}] {files[i][1].get('atchfileNm') or '이름 없음'} "
                                  f"(ID {files[i][1].get('contentId')}, #{i + 1})",
        )
        api, item = files[index]
        if st.button("첨부파일 가져오기"):
            try:
                path = attach.download(api, item)
                st.session_state["file"] = (item["atchfileUrl"], str(path))
            except attach.DownloadError as e:
                st.error(str(e))
        saved = st.session_state.get("file")
        if saved and saved[0] == item["atchfileUrl"]:
            path = Path(saved[1])
            st.success(f"저장 위치: {path}")
            st.download_button("브라우저로 내려받기", path.read_bytes(),
                               file_name=attach.safe_name(item.get("atchfileNm")))

con.close()
