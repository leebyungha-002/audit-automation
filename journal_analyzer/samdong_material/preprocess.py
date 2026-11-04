# -*- coding: utf-8 -*-
"""
journal_analyzer/samdong_material/preprocess.py
================================================
samdong_material 분개장 전용 전처리 (2026-11-04).

문제
----
ERP 분개장 내보내기가 5xx 계정에 '(제)', 8xx 계정에 '(판)'을 자동으로 덧붙인다.
 - 계정 마스터 이름에 이미 접미어가 있는 계정 → 두 번 붙음 ('광고선전비(판)(판)',
   당기 80개 중 36개, 전기 33개)
 - 마스터 이름에 접미어가 없는 계정 → 분개장에만 붙음 ('감가상각비(판)' vs 원장
   '63_감가상각비(81800)', 인력개발비·접대비(기업업무추진비)도 동일)
계정별원장 시트명은 마스터 이름 그대로라, 4번 TB_TO_GL_TEST 등 분개장↔원장 계정
매칭이 깨진다.

처리
----
1) 계정코드 기준 원장명 맞춤: data/current 계정별원장 xlsx의 시트명 'n_계정명(코드)'에서
   코드→계정명 표를 만들고(원장 코드 81800 = 분개장 코드 818 × 100), 분개장 계정명을
   원장 계정명으로 바꾼다. 감가상각비(제)/(판)은 원장 이름이 둘 다 '감가상각비'라
   하나로 합쳐진다(사용자 결정, 2026-11-04).
2) 원장에서 코드를 못 찾은 계정은 끝에 반복된 괄호 접미어만 1개로 줄인다
   ('광고선전비(판)(판)' -> '광고선전비(판)').
main_analyzer는 _preprocess_df(표준 컬럼명 정리) 전에 이 함수를 부르므로 원본 컬럼명
'계정과목'/'계정과목코드'를 기준으로 처리한다(표준명 '계정명'도 대비).
"""
import os
import re

import openpyxl
import pandas as pd

_COMPANY_DIR = os.path.dirname(os.path.abspath(__file__))
_CURRENT_DIR = os.path.join(_COMPANY_DIR, 'data', 'current')

_DUP_SUFFIX_RE = re.compile(r'(\([^()]+\))\1+$')
_SHEET_RE = re.compile(r'^\d+_(.+)\((\d+)\)$')
_ACCOUNT_COLS = ('계정과목', '계정명')
_CODE_COL = '계정과목코드'


def _dedupe_suffix(name):
    if not isinstance(name, str):
        return name
    return _DUP_SUFFIX_RE.sub(r'\1', name.strip())


def _ledger_code_names() -> dict:
    """원장 시트명에서 {분개장 계정코드(int): 원장 계정명}. 원장이 없으면 빈 딕셔너리."""
    if not os.path.isdir(_CURRENT_DIR):
        return {}
    out = {}
    for f in sorted(os.listdir(_CURRENT_DIR)):
        if '계정별원장' not in f or not f.endswith('.xlsx') or f.startswith('~$'):
            continue
        wb = openpyxl.load_workbook(os.path.join(_CURRENT_DIR, f), read_only=True)
        try:
            for sn in wb.sheetnames:
                m = _SHEET_RE.match(sn.strip())
                if not m:
                    continue
                code = int(m.group(2))
                if code % 100 == 0:
                    code //= 100
                out[code] = m.group(1).strip()
        finally:
            wb.close()
    return out


def _to_code(v):
    try:
        return int(float(v))
    except (TypeError, ValueError):
        return None


def preprocess(df: pd.DataFrame) -> pd.DataFrame:
    code_names = _ledger_code_names() if _CODE_COL in df.columns else {}
    codes = df[_CODE_COL].map(_to_code) if code_names else None

    for col in _ACCOUNT_COLS:
        if col not in df.columns:
            continue
        before = df[col].copy()
        cleaned = df[col].map(_dedupe_suffix)
        if code_names:
            ledger_name = codes.map(code_names)
            cleaned = ledger_name.where(ledger_name.notna(), cleaned)
        df[col] = cleaned

        changed = before.astype(str) != df[col].astype(str)
        if changed.any():
            pairs = (pd.DataFrame({'원본': before[changed], '변경': df.loc[changed, col]})
                     .drop_duplicates().sort_values('원본'))
            print(f'  [samdong_material] 계정명 정리({col}): {len(pairs)}개 이름, '
                  f'{int(changed.sum()):,}행 (원장 코드표 {len(code_names)}개 사용)')
            for _, r in pairs.iterrows():
                print(f'     {r["원본"]} -> {r["변경"]}')
    return df
