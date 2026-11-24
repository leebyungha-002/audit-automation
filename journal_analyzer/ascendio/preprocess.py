# -*- coding: utf-8 -*-
"""
journal_analyzer/ascendio/preprocess.py
=======================================
ascendio 회사 전용 분개장 전처리 모듈 (2026-11-24).

호출 시점: main_analyzer.py 의 load_data() 직후, 공통 표준 전처리(_preprocess_df) 이전.

원본 분개장 형식 (승인전표 조회 엑셀)
  승인일 / 승인번호 / 차변계정과목 / 대변계정과목 / 차변금액 / 대변금액 / 계정과목명 / 거래처명 / 적요 ...
  - 마지막 행이 '합계' 행(승인일='합계') → 그대로 두면 차·대변이 두 배로 집계됨
  - '차변계정과목'·'대변계정과목'은 "[8120000]여비교통비" 형태의 텍스트인데,
    공통 엔진은 이름에 '차변'/'대변'이 들어간 첫 컬럼을 금액으로 잡으므로 금액이 전부 0이 됨
  - 승인번호는 일자별로 다시 1부터 시작 → 공통 엔진이 전표그룹키(일자+번호)로 묶으므로 별도 처리 불필요
"""

import pandas as pd

COLUMN_MAP = {
    '승인일':     '전표일자',
    '승인번호':   '전표번호',
    '계정과목명': '계정명',
    '차변금액':   '차변',
    '대변금액':   '대변',
}

# 금액 컬럼으로 오인되지 않도록 '차변'/'대변'이 빠진 이름으로 보존
RENAME_SIDE_ACCOUNT = {
    '차변계정과목': '차측계정(원본)',
    '대변계정과목': '대측계정(원본)',
}


def _clean_journal_id(v):
    if v is None or (isinstance(v, float) and v != v):
        return None
    try:
        f = float(v)
        return str(int(f)) if f.is_integer() else str(v).strip()
    except (TypeError, ValueError):
        return str(v).strip()


def preprocess(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    # ascendio 형식이 아니면(이미 표준 컬럼이거나 다른 파일) 손대지 않음
    if '승인일' not in df.columns and '전표일자' in df.columns:
        return df

    df = df.rename(columns={**RENAME_SIDE_ACCOUNT, **COLUMN_MAP})

    # 합계 행 제거 (승인일='합계' 또는 일자 없음)
    before = len(df)
    date_txt = df['전표일자'].astype(str).str.strip()
    df = df[(date_txt != '합계') & df['전표일자'].notna() & (date_txt != '')]
    removed = before - len(df)

    df['전표번호'] = df['전표번호'].map(_clean_journal_id)
    for c in ('차변', '대변'):
        df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0)

    # 검증: 차대 합계 일치
    dr, cr = df['차변'].sum(), df['대변'].sum()
    print(f'  [ascendio 전처리] 합계행 {removed}건 제거, {len(df)}행 / 차변 {dr:,.0f} / 대변 {cr:,.0f}'
          + ('' if abs(dr - cr) < 1 else '  ⚠️ 차대 불일치'))
    return df.reset_index(drop=True)
