# -*- coding: utf-8 -*-
"""
journal_analyzer/daesun/preprocess.py
======================================
daesun(대선제분(주)) 분개장 전용 전처리.

문제
----
원본 컬럼명이 main_analyzer._normalize_date_journal_columns() 의 기본 별칭
목록과 일부 다르다 (2026-08-27 확인, 컬럼: 회계일자/NO/전표번호/라인번호/계정명/
차변/대변/적요/거래처코드/거래처/작성부서/작성사원/회계단위/전표유형/전표구분/승인자).

  - '회계일자' : 기본 별칭 목록(일자/전표일자/전표일/거래일자/적요일자)에
    없어 자동 매핑되지 않음 → 명시적으로 '전표일자'로 변경 필요.
    (전표일자가 없으면 COL_JOURNAL_KEY, 일자차이분석, 데이터개요 등
     날짜 기반 로직이 전부 깨짐)
  - '작성사원' : analyze_employee_summary() 의 자동 탐지 키워드
    ('사원명'/'작성자'/'사용자'/'User'/'Employee')와 매칭되지 않음
    → 명시적으로 '사원명'으로 변경 필요.
  - '거래처'/'계정명'은 이미 표준 별칭 목록에 포함되어 있어 그대로 둠.

전표번호(예: FI2025010100001)는 kyungnam/sejoong과 달리 날짜별 재사용
없이 고유한 것으로 보여(2026-08-27 샘플 1건 확인), 별도 전표그룹키
재구성 없이 main_analyzer 기본 로직(전표일자+전표번호)을 그대로 사용한다.
데이터가 더 들어오면 이 가정을 재검증할 것.

main_analyzer.py 호출 규약
--------------------------
함수 시그니처 :  preprocess(df: pd.DataFrame) -> pd.DataFrame
"""

import pandas as pd

COLUMN_MAP: dict[str, str] = {
    '회계일자': '전표일자',
    '작성사원': '사원명',
}


def preprocess(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    rename_dict = {}
    already_in_df = set(df.columns)
    for orig, target in COLUMN_MAP.items():
        if orig not in already_in_df:
            continue
        if target in already_in_df and target != orig:
            continue
        rename_dict[orig] = target

    df = df.rename(columns=rename_dict)
    print(f'  [daesun/preprocess] 컬럼 매핑 적용: {rename_dict}')
    return df
