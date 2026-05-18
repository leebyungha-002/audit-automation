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
  1. _map_columns()  : 비표준 컬럼명 → 표준 컬럼명 매핑
  2. _build_pk()     : 전표번호 + 순번 결합 → 고유 식별키(PK) 생성
  3. preprocess()    : 위 두 단계를 순서대로 호출하는 공개 진입점

커스터마이징 방법
-----------------
kyungnam 엑셀 파일의 실제 컬럼명을 확인한 후
아래 COLUMN_MAP 의 좌측(원본)을 수정한다.
우측(표준명)은 main_analyzer.py 의 COL_* 상수와 일치해야 한다.
"""

import sys
import pandas as pd

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')

# =============================================================================
# 1. 헤더 매핑 테이블  (실제 파일 기준: 경남제약 분개장)
#
#  확인된 원본 컬럼 (9개)
#  ┌────────────┬──────────────────────────────────────────┐
#  │ 원본 컬럼  │ 비고                                      │
#  ├────────────┼──────────────────────────────────────────┤
#  │ 전표일자   │ 표준명 동일 → 매핑 불필요                  │
#  │ 전표번호   │ 표준명 동일 → 매핑 불필요                  │
#  │ 순번       │ PK 생성에 사용 (표준 컬럼 아님)             │
#  │ 계정코드   │ 분석에 미사용 (그대로 유지)                 │
#  │ 계정명     │ 표준명 동일 → 매핑 불필요                  │
#  │ 전표구분   │ 분석에 미사용 (그대로 유지)                 │
#  │ 차변금액   │ ★ '차변'으로 매핑 필요                     │
#  │ 대변금액   │ ★ '대변'으로 매핑 필요                     │
#  │ 적요       │ 표준명 동일 → 매핑 불필요                  │
#  └────────────┴──────────────────────────────────────────┘
#
#  ※ 누락 컬럼: 거래처명, 사원명
#     → 거래처명이 없어 거래처비교·특수관계자 등 분석은 빈 결과 반환
#     → preprocess() 에서 빈 컬럼을 자동 생성하여 분석 오류를 방지함
# =============================================================================
COLUMN_MAP: dict[str, str] = {
    '차변금액': '차변',
    '대변금액': '대변',
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
        # 목표 컬럼이 이미 존재하고 자기 자신이 아닌 경우 → 충돌 방지를 위해 skip
        if target in already_in_df and target != orig:
            continue
        if orig != target:
            rename_dict[orig] = target

    return df.rename(columns=rename_dict)


def _build_pk(df: pd.DataFrame, *, zero_pad_jid: int = 8, zero_pad_seq: int = 3) -> pd.DataFrame:
    """
    '전표번호' 와 '순번' 을 결합하여 고유 식별키(PK) 컬럼을 생성한다.

    생성 규칙
    ---------
    두 컬럼 모두 존재  → '{전표번호:08d}-{순번:03d}'  예) '00002025-003'
    전표번호만 존재    → '{전표번호}' 그대로 사용
    둘 다 없음         → 행 인덱스 문자열로 대체 (경고 출력)

    Parameters
    ----------
    zero_pad_jid : 전표번호 제로패딩 자릿수 (기본 8)
    zero_pad_seq : 순번 제로패딩 자릿수 (기본 3)
    """
    df = df.copy()
    has_jid = '전표번호' in df.columns
    has_seq = '순번'     in df.columns

    if has_jid and has_seq:
        jid = (df['전표번호']
               .fillna('')
               .astype(str)
               .str.strip()
               .str.zfill(zero_pad_jid))
        seq = (df['순번']
               .fillna('')
               .astype(str)
               .str.strip()
               .str.zfill(zero_pad_seq))
        df[PK_COL] = jid + '-' + seq

    elif has_jid:
        df[PK_COL] = df['전표번호'].fillna('').astype(str).str.strip()

    else:
        print(f'  [kyungnam/preprocess] ⚠ 전표번호·순번 컬럼을 찾지 못했습니다. '
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

    Parameters
    ----------
    df : load_data() 가 반환한 원본(raw) 분개장 DataFrame

    Returns
    -------
    pd.DataFrame : 컬럼명이 표준화되고 PK 컬럼이 추가된 DataFrame
    """
    print('  [kyungnam/preprocess] 컬럼 매핑 시작')

    # Step 1 — 컬럼명 표준화 (차변금액→차변, 대변금액→대변)
    df = _map_columns(df)

    # Step 2 — 누락 표준 컬럼 보정
    # 거래처명·사원명이 없으면 빈 문자열 컬럼으로 생성하여 분석 함수 오류 방지
    MISSING_FILL = {'거래처명': '', '사원명': ''}
    for col, fill in MISSING_FILL.items():
        if col not in df.columns:
            df[col] = fill
            print(f'  [kyungnam/preprocess] "{col}" 컬럼 없음 → 빈 컬럼 생성')

    print(f'  [kyungnam/preprocess] 최종 컬럼: {list(df.columns)}')

    # Step 3 — 고유 식별키 생성 (전표번호 + 순번)
    df = _build_pk(df)
    print(f'  [kyungnam/preprocess] PK 샘플: {df[PK_COL].head(3).tolist()}')

    return df
