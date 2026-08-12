# -*- coding: utf-8 -*-
"""
journal_analyzer/graphy/preprocess.py
======================================
graphy 회사 전용 분개장 전처리 모듈.

호출 시점
---------
main_analyzer.py 의 load_data() 완료 직후,
표준 전처리(_preprocess_df()) 호출 이전에 자동으로 실행된다.

처리 흐름
---------
  1. _map_columns()        : '날짜' -> '전표일자' 매핑
                              (전표번호/차변/대변/거래처명/적요/계정과목은
                               이미 표준 컬럼명과 일치하거나 main_analyzer의
                               자동 매핑 목록에 있어 그대로 인식됨)
  2. _strip_account_code()  : 계정과목 앞의 '[코드]' 접두어 제거
                              (예: '[25402]국민연금_예수금' -> '국민연금_예수금')
                              — 코드가 붙은 채로 둬도 contains 폴백으로 매칭은
                                되지만, _account_match_flexible() 의 접두어
                                오매칭 위험(2026-08-12 발견)을 줄이기 위함
  3. preprocess()            : 위 두 단계를 순서대로 호출하는 공개 진입점
"""

import re
import sys
import pandas as pd

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')

# =============================================================================
# 1. 헤더 매핑 테이블  (실제 파일 기준: graphy 분개장, 14개 컬럼)
#
#  원본 컬럼   -> 표준 컬럼명   비고
#  날짜        -> 전표일자      main_analyzer 자동 매핑 목록에 '날짜'가 없어 명시 매핑 필요
#  전표번호    -> (유지)        표준명 동일
#  계정과목    -> (유지)        main_analyzer fallback 이 처리 (COLUMN_MAP 불필요)
#  차변/대변   -> (유지)        표준명 동일
#  거래처명    -> (유지)        표준명 동일
#  적요        -> (유지)        표준명 동일
#  코드/년/월/일/번호/코드.1/PJT명  분석 미사용, 그대로 유지
# =============================================================================
COLUMN_MAP: dict[str, str] = {
    '날짜': '전표일자',
}

# 계정과목 앞에 붙은 '[코드]' 접두어 제거용 정규식
_ACCOUNT_CODE_PREFIX = re.compile(r'^\[[^\]]*\]\s*')


# =============================================================================
# 2. 내부 헬퍼 함수
# =============================================================================

def _map_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    COLUMN_MAP 을 기준으로 원본 컬럼명을 표준명으로 변경한다.

    규칙
    ----
    - 매핑 대상이 아닌 컬럼은 그대로 유지한다.
    - 이미 표준명이 DataFrame 에 존재하는 경우 중복 rename 을 건너뛴다.
    """
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]   # 앞뒤 공백 제거

    rename_dict: dict[str, str] = {}
    already_in_df = set(df.columns)

    for orig, target in COLUMN_MAP.items():
        if orig not in already_in_df:
            continue
        if target in already_in_df and target != orig:
            continue
        if orig != target:
            rename_dict[orig] = target

    return df.rename(columns=rename_dict)


def _strip_account_code(df: pd.DataFrame) -> pd.DataFrame:
    """
    '계정과목' 컬럼 값 앞의 '[코드]' 접두어를 제거한다.
    (예: '[25402]국민연금_예수금' -> '국민연금_예수금')
    """
    df = df.copy()
    for col in ('계정과목', '계정명'):
        if col in df.columns:
            df[col] = (df[col]
                       .fillna('')
                       .astype(str)
                       .str.replace(_ACCOUNT_CODE_PREFIX, '', regex=True)
                       .str.strip())
    return df


# =============================================================================
# 3. 공개 진입점 — main_analyzer.py 가 동적으로 호출하는 함수
# =============================================================================

def preprocess(df: pd.DataFrame) -> pd.DataFrame:
    """
    graphy 원본 분개장 DataFrame 을 표준 형식으로 전처리하여 반환한다.

    main_analyzer.py 호출 규약
    --------------------------
    함수 시그니처 :  preprocess(df: pd.DataFrame) -> pd.DataFrame
    모듈 내 반드시 이 이름으로 정의되어야 한다.
    """
    print('  [graphy/preprocess] 컬럼 매핑 시작')

    # Step 1 — 컬럼명 표준화 (날짜 -> 전표일자)
    df = _map_columns(df)
    print(f'  [graphy/preprocess] 최종 컬럼: {list(df.columns)}')

    # Step 2 — 계정과목 코드 접두어 제거
    df = _strip_account_code(df)

    return df
