# -*- coding: utf-8 -*-
"""
journal_analyzer/samdong_material/preprocess.py
================================================
samdong_material 분개장 전용 전처리 (2026-11-04).

문제
----
ERP에서 내려받은 분개장의 계정과목명 끝에 판관/제조 구분 접미어가 두 번 붙어 있다
(예: '광고선전비(판)(판)', '급여(제)(제)' — 당기 80개 중 36개, 전기 33개 계정).
계정별원장 시트명은 '38_급여(제)(50300)'처럼 한 번만 붙어 있어 4번 TB_TO_GL_TEST 등
분개장↔원장 계정 매칭이 깨진다. '감가상각비(판)', '접대비(기업업무추진비)(판)'처럼
정상인 계정도 섞여 있다.

처리
----
계정명 끝에 같은 괄호 접미어가 연속 반복되면 1개로 줄인다:
  '광고선전비(판)(판)' -> '광고선전비(판)',  '접대비(기업업무추진비)(판)' -> 그대로
main_analyzer는 _preprocess_df(표준 컬럼명 정리) 전에 이 함수를 부르므로 원본 컬럼명
'계정과목'과 표준 컬럼명 '계정명'을 모두 처리한다.
"""
import re

import pandas as pd

_DUP_SUFFIX_RE = re.compile(r'(\([^()]+\))\1+$')
_ACCOUNT_COLS = ('계정과목', '계정명')


def _dedupe_suffix(name):
    if not isinstance(name, str):
        return name
    s = name.strip()
    return _DUP_SUFFIX_RE.sub(r'\1', s)


def preprocess(df: pd.DataFrame) -> pd.DataFrame:
    for col in _ACCOUNT_COLS:
        if col not in df.columns:
            continue
        before = df[col].copy()
        df[col] = df[col].map(_dedupe_suffix)
        changed = before[before.astype(str) != df[col].astype(str)]
        if not changed.empty:
            print(f'  [samdong_material] 계정명 중복 접미어 정리({col}): '
                  f'{changed.nunique()}개 계정, {len(changed):,}행')
    return df
