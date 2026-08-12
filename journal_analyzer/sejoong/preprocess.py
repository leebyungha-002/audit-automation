# -*- coding: utf-8 -*-
"""
journal_analyzer/sejoong/preprocess.py
=======================================
sejoong(더존 ERP) 분개장 전용 전처리.

문제
----
원본 CSV의 '전표번호' 컬럼이 이미 과학적 표기법 문자열(예: "2.02601E+12")로
정밀도가 소실된 채로 export되어 있다(2026-08-12 발견) — 418,740행 중 서로 다른
전표번호가 6종류로 뭉개진 상태라 전표번호로는 "같은 전표"를 식별할 수 없다.
날짜와 결합해도 복구 불가능한 수준(회사 담당자가 더존에서 내려받아 저장하는
과정에서 이미 깨진 것으로 추정 — 업로드 전 수기로 보정하던 관행이 이 문제 때문).

해결
----
같은 파일의 '항번호'(전표 내 분개행 순번) 컬럼이 새 전표가 시작될 때마다 1로
리셋되는 패턴을 이용해 전표 경계를 재구성한다.
검증(2026-08-12): 이 방식으로 재구성한 57,291개 전표 전부 차변합계=대변합계로
완전히 대차평형됨을 확인 — 전표번호 없이도 정확한 전표 매칭이 가능함.

main_analyzer.py 호출 규약
--------------------------
함수 시그니처 :  preprocess(df: pd.DataFrame) -> pd.DataFrame
main_analyzer._preprocess_df() 가 이 함수 실행 이후 호출되며, 이미
'전표그룹키' 컬럼이 존재하면 main_analyzer 쪽 기본 생성 로직(전표일자+전표번호)은
건드리지 않고 이 값을 그대로 존중한다.
"""

import pandas as pd

# main_analyzer.COL_JOURNAL_KEY 와 동일한 값 (순환 임포트 방지를 위해 문자열 직접 사용)
JOURNAL_KEY_COL = '전표그룹키'


def preprocess(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    if '항번호' not in df.columns or '전표일자' not in df.columns:
        print('  [sejoong/preprocess] 항번호 또는 전표일자 컬럼 없음 → 전표그룹키 재구성 건너뜀')
        return df

    hang = pd.to_numeric(df['항번호'], errors='coerce')
    voucher_seq = (hang == 1).cumsum()

    date_num = pd.to_numeric(df['전표일자'], errors='coerce')
    if date_num.notna().mean() > 0.5:
        date_str = date_num.fillna(0).astype('int64').astype(str)
    else:
        date_str = df['전표일자'].astype(str).str.strip()

    df[JOURNAL_KEY_COL] = date_str + '_V' + voucher_seq.astype(str)
    print(f'  [sejoong/preprocess] 항번호 리셋 기반 전표그룹키 재구성 완료 '
          f'(전표번호 손상 대응, 재구성된 전표 수: {int(voucher_seq.max())})')

    return df
