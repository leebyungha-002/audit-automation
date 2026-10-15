"""
journal_analyzer/live_pure/preprocess.py
live_pure 분개장 전용 전처리 — 계정명 동일·계정코드 상이 계정 분리

배경 (2026-10-15)
-----------------
live_pure 원장은 서로 다른 계정이 같은 계정명을 쓰는 경우가 있다.
분석 엔진은 계정코드 없이 계정명으로만 매칭하므로 그대로 두면 합쳐져 버린다.
  - 퇴직급여(판) 806(퇴직충당금설정, DB형) / 808(퇴직연금 DC확정기여형)
    → 25번 손익월별분석 등에서 한 줄로 합산됨
  - 감가상각누계액 213(비품) / 220(리스개량자산)
    → 27번 유형자산 롤포워드에서 두 자산의 누계액이 섞임
계정코드 기준으로 계정명을 구분되게 바꿔 당기·전기 모두 동일하게 적용한다.

퇴직급여는 '-DB'/'-DC'를 뒤에 붙여서, task_list에 '퇴직급여(판)'만 적으면
startswith 매칭으로 두 계정 합계가, '퇴직급여(판)-DB'처럼 적으면 해당 계정만 잡힌다.

main_analyzer.py 호출 규약
--------------------------
함수 시그니처 :  preprocess(df: pd.DataFrame) -> pd.DataFrame
main_analyzer._preprocess_df()(컬럼 표준화) 이전에 호출되므로 원본 컬럼명
('Code', '계정과목')을 기준으로 처리한다.
"""

import pandas as pd

CODE_COL = 'Code'
ACCOUNT_COLS = ('계정과목', '계정명')

# 계정코드 → 분리 후 계정명
CODE_RENAME = {
    '806': '퇴직급여(판)-DB',
    '808': '퇴직급여(판)-DC',
    '213': '감가상각누계액(비품)',
    '220': '감가상각누계액(리스개량자산)',
}


def preprocess(df: pd.DataFrame) -> pd.DataFrame:
    acct_col = next((c for c in ACCOUNT_COLS if c in df.columns), None)
    if CODE_COL not in df.columns or acct_col is None:
        print('  [live_pure/preprocess] Code 또는 계정과목 컬럼 없음 → 계정명 분리 건너뜀')
        return df

    # Code가 float(806.0)으로 읽히는 경우가 있어 정수 문자열로 통일
    codes = pd.to_numeric(df[CODE_COL], errors='coerce').astype('Int64').astype(str)
    for code, new_name in CODE_RENAME.items():
        mask = codes == code
        if mask.any():
            df.loc[mask, acct_col] = new_name
            print(f'  [live_pure/preprocess] 계정코드 {code} → "{new_name}" ({mask.sum()}행)')
    return df
