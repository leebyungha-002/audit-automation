"""금감원 회계업무자료 Open API 클라이언트"""
import html
import re

import requests

BASE = "https://www.fss.or.kr/fss/kr/openApi/api/"
ENDPOINTS = {
    "주석공시": "illustDisc.jsp",  # 안내문에서 확인
    # 아래는 홈페이지 OPEN API '상세 및 테스트' 화면(2026-11-28 확인)의 요청URL.
    # 요청변수는 주석공시와 동일(apiType/startDate/endDate/authKey).
    # 실제 키로 호출 검증 전이므로, 실패하면 해당 API만 건너뛰고 오류를 표시한다.
    "K-IFRS 질의회신": "kifrsQa.jsp",
    "일반기준 질의회신": "kgaapQa.jsp",
    "내부회계관리제도": "icfrData.jsp",
    "심사·감리지적사례": "inspExmpl.jsp",
    "외부감사 업무자료": "audData.jsp",
    "외부감사 FAQ": "audFaq.jsp",
}

# 제목·본문 외에 검색에 포함할 원본 항목 (관련 기준서, 쟁점 분야)
EXTRA_SEARCH_FIELDS = ("relStand", "issRealm")


class FssApiError(Exception):
    pass


def clean_html(text: str) -> str:
    """이스케이프된 HTML을 일반 텍스트로 변환"""
    text = html.unescape(text or "")              # &lt;p&gt; -> <p>
    text = re.sub(r"<(br|/p)\b[^>]*>", "\n", text, flags=re.I)
    text = re.sub(r"</?[a-zA-Z][^>]*>", "", text)  # 태그 제거
    text = html.unescape(text).replace("\xa0", " ")  # &nbsp;
    return re.sub(r"\n\s*\n+", "\n", text).strip()


def parse_response(payload: dict) -> list[dict]:
    body = payload.get("reponse") or payload.get("response")
    if not isinstance(body, dict):
        raise FssApiError(f"예상과 다른 응답 구조: {list(payload)}")
    if str(body.get("resultCode")) != "1":
        raise FssApiError(f"{body.get('resultCode')}: {body.get('resultMsg')}")
    items = body.get("result") or []
    if body.get("resultCnt") is not None and int(body["resultCnt"]) != len(items):
        # 감사 목적상 누락 가능성은 조용히 넘기지 않는다
        raise FssApiError(f"건수 불일치: resultCnt={body['resultCnt']}, 수신={len(items)}")
    return items


def fetch(api_name: str, start: str, end: str, key: str) -> list[dict]:
    params = {"apiType": "json", "startDate": start, "endDate": end, "authKey": key}
    try:
        res = requests.get(BASE + ENDPOINTS[api_name], params=params, timeout=30)
        res.raise_for_status()
        payload = res.json()
    except (requests.RequestException, ValueError) as e:
        # 원래 예외에는 인증키가 든 URL이 포함될 수 있어 연결을 끊는다
        raise FssApiError(f"호출 실패: {type(e).__name__}") from None
    return parse_response(payload)
