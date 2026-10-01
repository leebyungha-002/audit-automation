"""분개장 집계 신호 (확장 지점, 1차 개발 범위 밖).

나중에 journal_analyzer의 분개장 집계 요약(기말 집중 전표, 수기 전표 비중, 특수관계자 거래 등)을
읽어 표준 계정 분류별 "위험 발현 가능성" 문구를 돌려주도록 구현한다.
report 모듈은 이 함수가 돌려준 분류에만 해당 열을 채우며, 빈 dict면 열 자체를 만들지 않는다.
"""
from .db import Store


def load_signals(store: Store, company: str) -> dict[str, str]:
    """{표준 분류: 위험 발현 가능성 설명}. 아직 구현하지 않았으므로 빈 dict."""
    return {}
