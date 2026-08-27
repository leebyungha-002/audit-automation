# -*- coding: utf-8 -*-
"""
journal_analyzer/samdong/preprocess.py
=======================================
samdong 분개장 전용 전처리.

원본 포맷 (2026-08-27 확인, 더존 계열 ERP 추정)
------------------------------------------------
컬럼: 회계단위/전표관리단위/전표기표번호/전표승인번호/행/계정과목/비용구분/
차변금액/대변금액/적요/외화차변금액/외화대변금액/통화/환율/증빙/기표자/
귀속부서/활동센터/관리항목1~8

문제
----
1. 전용 날짜 컬럼(전표일자/회계일자)이 없다. 회계일자는 관리항목 컬럼 중
   하나에 문자열로 들어있는데, 그 위치가 행(계정)마다 다르다
   (예: 여비교통비-식대 행은 관리항목2, 미지급금-카드 행은 관리항목3에
   '2025-01-01'이 위치 — 관리항목은 계정과목별로 의미가 달라지는
   가변 슬롯이라 고정 컬럼으로 신뢰할 수 없음).
2. '전표번호'에 해당하는 컬럼이 두 개다.
   - 전표기표번호: 라인(행) 단위로 유일 (예: "11-11-20250101-0002-001",
     차/대변 각 줄마다 접미어(-001,-002...)가 달라짐) → 상대계정 매칭에
     쓰면 차변/대변이 서로 다른 그룹으로 잡혀 매칭이 안 됨.
   - 전표승인번호: 같은 전표의 차/대변 라인이 공유 (예: "11-20250101-0002")
     → 상대계정분석(8번 메뉴) 등 "같은 전표에 속한 거래" 매칭에 적합
     (사용자 확인, 2026-08-27).
3. 거래처명 전용 컬럼이 없음 → 관리항목1을 거래처명으로 간주 (사용자 확인,
   2026-08-27). 단, 확보된 샘플 4행(여비교통비-식대/미지급금-카드, 비용·
   미지급금 라인)에서는 관리항목1이 '신용카드'/카드번호처럼 거래처명이
   아닌 값이었다 — 관리항목 슬롯은 계정과목별로 의미가 달라지므로, 매출
   계정(외상매출금 등)에서 관리항목1이 실제 거래처명이 맞는지는 실제
   매출 관련 데이터로 재확인 필요.

해결
----
전표승인번호(예: "11-20250101-0002")를 기준으로:
  - 전표번호 / 전표그룹키 컬럼을 전표승인번호 값 그대로 생성
    (차/대변 라인이 자동으로 같은 그룹으로 묶임 → 상대계정분석 가능)
  - 문자열 중 8자리 숫자(YYYYMMDD, 예: "20250101")를 정규식으로 추출해
    전표일자 컬럼 생성. main_analyzer._preprocess_df() 가 이 문자열을
    표준 날짜 파싱 로직(YYYYMMDD 포맷)으로 이어받아 datetime으로
    변환하므로, 여기서는 파싱된 datetime이 아니라 원본 문자열째로 넘긴다
    (미리 datetime으로 바꿔두면 _preprocess_df 의 재파싱 로직이 numeric
    변환을 시도하다가 깨짐 — 다른 회사 preprocess.py들과 동일 관례).

main_analyzer.py 호출 규약
--------------------------
함수 시그니처 :  preprocess(df: pd.DataFrame) -> pd.DataFrame
"""

import re
import pandas as pd

# main_analyzer.COL_JOURNAL_ID / COL_JOURNAL_KEY 와 동일한 값
# (순환 임포트 방지를 위해 문자열 직접 사용)
JOURNAL_ID_COL  = '전표번호'
JOURNAL_KEY_COL = '전표그룹키'
DATE_COL        = '전표일자'

_DATE_RE = re.compile(r'(20\d{6})')


def preprocess(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    if '전표승인번호' not in df.columns:
        print('  [samdong/preprocess] 전표승인번호 컬럼 없음 → 전처리 건너뜀')
        return df

    approval = df['전표승인번호'].astype(str).str.strip()

    # 전표승인번호를 전표번호/전표그룹키로 그대로 사용
    # (같은 전표의 차/대변 라인이 이 값을 공유하므로 상대계정분석에 바로 적합)
    df[JOURNAL_ID_COL]  = approval
    df[JOURNAL_KEY_COL] = approval

    # 전표승인번호 안의 YYYYMMDD 추출 → 전표일자 (문자열째로 넘겨 표준 파싱에 위임)
    date_str = approval.str.extract(_DATE_RE, expand=False)
    df[DATE_COL] = date_str

    # 거래처명 전용 컬럼이 없는 회사 → 관리항목1을 거래처명으로 간주 (사용자 확인, 2026-08-27)
    if '거래처명' not in df.columns:
        if '관리항목1' in df.columns:
            df['거래처명'] = df['관리항목1']
        else:
            df['거래처명'] = ''
    # 기표자 → 사원명 (사원별집계 메뉴 인식용)
    if '사원명' not in df.columns and '기표자' in df.columns:
        df = df.rename(columns={'기표자': '사원명'})

    n_ok = date_str.notna().sum()
    print(f'  [samdong/preprocess] 전표승인번호 기반 전표번호/전표그룹키/전표일자 생성 완료 '
          f'(날짜 추출 성공 {n_ok}/{len(df)}행)')
    return df
