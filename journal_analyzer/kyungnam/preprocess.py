# -*- coding: utf-8 -*-
"""
journal_analyzer/kyungnam/preprocess.py
=======================================
kyungnam 회사 전용 분개장 전처리 모듈.

호출 시점
---------
main_analyzer.py 의 load_data() 완료 직후,
표준 전처리(_preprocess_df()) 호출 이전에 자동으로 실행된다.

처리 흐름
---------
  1. _map_columns()  : 비표준 컬럼명 -> 표준 컬럼명 매핑
  2. _build_pk()     : 전표번호 결합 -> 고유 식별키(PK) 생성
  3. preprocess()    : 위 두 단계를 순서대로 호출하는 공개 진입점
"""

import sys
import pandas as pd

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')

# =============================================================================
# 1. 헤더 매핑 테이블  (실제 파일 기준: 경남제약 분개장, 20개 컬럼)
#
#  원본 컬럼      표준 컬럼명    비고
#  결의일자    -> 전표일자       날짜 기반 분析의 기준일
#  결의번호    -> 전표번호       전표 식별 키
#  사원        -> 사원명         사원별집계 분析용
#  차변금액    -> 차변           금액 분析 기준
#  대변금액    -> 대변           금액 분析 기준
#  회계일자    -> 등록일자       일자차이분析: 결의 -> 회계 지연 탐지
#  계정과목    -> 계정명         main_analyzer fallback 이 처리 (COLUMN_MAP 불필요)
#  거래처명    -> (유지)         표준명 동일
#  적요        -> (유지)         표준명 동일
#  부서/차대구분/승인요청일 등   분析 미사용, 그대로 유지
# =============================================================================
COLUMN_MAP: dict[str, str] = {
    '결의일자': '전표일자',
    '결의번호': '전표번호',
    '차변금액': '차변',
    '대변금액': '대변',
    '사원':     '사원명',
    '회계일자': '등록일자',   # 일자차이분析: 결의일자 vs 회계일자 지연 탐지
}

# PK 컬럼명 (main_analyzer 가 참조할 수 있도록 상수로 노출)
PK_COL = 'PK'


# =============================================================================
# 2. 내부 헬퍼 함수
# =============================================================================

def _map_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    COLUMN_MAP 을 기준으로 원본 컬럼명을 표준명으로 변경한다.

    규칙
    ----
    - 매핑 대상이 아닌 컬럼은 그대로 유지한다.
    - 이미 표준명이 DataFrame 에 존재하는 경우 중복 rename 을 건너뛴다
      (원본 데이터가 이미 혼합 컬럼을 갖는 경우 방어).
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


def _build_pk(df: pd.DataFrame, *, zero_pad_jid: int = 8) -> pd.DataFrame:
    """
    '전표번호' 를 기반으로 고유 식별키(PK) 컬럼을 생성한다.
    (경남제약 파일에는 순번 컬럼이 없으므로 전표번호 단독 사용)

    전표번호 없음 -> 행 인덱스 문자열로 대체 (경고 출력)
    """
    df = df.copy()
    has_jid = '전표번호' in df.columns

    if has_jid:
        df[PK_COL] = (df['전표번호']
                      .fillna('')
                      .astype(str)
                      .str.strip()
                      .str.zfill(zero_pad_jid))
    else:
        print(f'  [kyungnam/preprocess] 전표번호 컬럼을 찾지 못했습니다. '
              f'PK 를 행 인덱스({len(df)}건)로 대체합니다.')
        df[PK_COL] = df.reset_index(drop=True).index.astype(str)

    return df


# =============================================================================
# 3. 공개 진입점 — main_analyzer.py 가 동적으로 호출하는 함수
# =============================================================================

def preprocess(df: pd.DataFrame) -> pd.DataFrame:
    """
    kyungnam 원본 분개장 DataFrame 을 표준 형식으로 전처리하여 반환한다.

    main_analyzer.py 호출 규약
    --------------------------
    함수 시그니처 :  preprocess(df: pd.DataFrame) -> pd.DataFrame
    모듈 내 반드시 이 이름으로 정의되어야 한다.
    """
    print('  [kyungnam/preprocess] 컬럼 매핑 시작')

    # Step 1 — 컬럼명 표준화
    df = _map_columns(df)

    # Step 2 — 누락 표준 컬럼 보정 (분析 함수 오류 방지)
    MISSING_FILL = {'거래처명': '', '사원명': ''}
    for col, fill in MISSING_FILL.items():
        if col not in df.columns:
            df[col] = fill
            print(f'  [kyungnam/preprocess] "{col}" 컬럼 없음 -> 빈 컬럼 생성')

    print(f'  [kyungnam/preprocess] 최종 컬럼: {list(df.columns)}')

    # Step 3 — 고유 식별키 생성
    df = _build_pk(df)
    print(f'  [kyungnam/preprocess] PK 샘플: {df[PK_COL].head(3).tolist()}')

    return df
