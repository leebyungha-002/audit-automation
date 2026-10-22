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

입금전표 현금 라인 보충 (2026-10-22)
------------------------------------
전표구분='입금'인 전표는 ERP 출력에 대변 라인(특정금전신탁/당좌예금 인출 등)만
있고 상대편 현금 차변 라인이 빠져 있다 (당기 35건 1.837억, 전기 40건 확인 —
계정별원장도 동일하게 누락돼 원장상 현금 잔액이 음수). 이 상태로는 해당 전표의
차대가 맞지 않고 상대계정분석에서 상대계정이 비어 나오므로, 입금전표 중 차대가
맞지 않는 전표마다 (대변합계-차변합계) 금액의 '현금' 차변 라인을 1줄 추가한다.
전표구분='출금'은 반대로 차변 라인만 있을 것이므로 (차변합계-대변합계) 금액의
'현금' 대변 라인을 추가한다 (2026-10-22 기준 데이터에 출금전표는 없음 — 대비용).
적요 끝에 ' (현금 보충)'을 붙여 원본 라인과 구분한다.

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
    return _add_missing_cash_lines(df)


CASH_ACCOUNT = '현금'
CASH_NOTE_SUFFIX = ' (현금 보충)'


def _add_missing_cash_lines(df: pd.DataFrame) -> pd.DataFrame:
    """입금/출금 전표 중 차대 불일치분에 현금 라인을 추가한다.
    입금 → 현금 차변(대변-차변), 출금 → 현금 대변(차변-대변)."""
    need = {'전표구분', '전표번호', '계정명', '차변', '대변'}
    if not need.issubset(df.columns):
        return df

    kind = df['전표구분'].astype(str).str.strip()
    added = []
    # (전표구분, 현금이 들어갈 칸, 반대 칸)
    for label, cash_col, other_col in (('입금', '차변', '대변'), ('출금', '대변', '차변')):
        mask = kind == label
        if not mask.any():
            continue

        sub = df[mask].copy()
        sub['_cash'] = pd.to_numeric(sub[cash_col], errors='coerce').fillna(0)
        sub['_other'] = pd.to_numeric(sub[other_col], errors='coerce').fillna(0)
        # 당기/전기 파일이 합쳐진 상태라 '구분'(당기/전기)도 키에 포함
        keys = ['전표번호'] + (['구분'] if '구분' in sub.columns else [])
        sums = sub.groupby(keys)[['_cash', '_other']].sum()
        gap = (sums['_other'] - sums['_cash']).round(2)
        gap = gap[gap > 0]
        if gap.empty:
            continue

        # 전표의 첫 라인을 복사해 공통 정보(일자·NO·거래처·부서 등)를 유지
        first = (sub.drop(columns=['_cash', '_other'])
                 .groupby(keys, sort=False).head(1).set_index(keys))
        new = first.loc[gap.index].reset_index()
        new['계정명'] = CASH_ACCOUNT
        new[cash_col] = gap.values
        new[other_col] = 0
        if '라인번호' in new.columns:
            max_line = sub.groupby(keys)['라인번호'].max()
            new['라인번호'] = (pd.to_numeric(max_line.loc[gap.index], errors='coerce')
                             .fillna(0).astype(int).values + 1)
        if '적요' in new.columns:
            new['적요'] = new['적요'].fillna('').astype(str) + CASH_NOTE_SUFFIX
        added.append(new[df.columns])

        print(f'  [daesun/preprocess] {label}전표 현금 {cash_col} 라인 보충: '
              f'{len(new)}건, {gap.sum():,.0f}원')

    if not added:
        return df
    return pd.concat([df] + added, ignore_index=True)
