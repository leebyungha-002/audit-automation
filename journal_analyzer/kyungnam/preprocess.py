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

import os
import re
import sys

import openpyxl
import pandas as pd

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')

# =============================================================================
# 1. 헤더 매핑 테이블  (실제 파일 기준: 경남제약 분개장, 20개 컬럼)
#
#  원본 컬럼      표준 컬럼명    비고
#  결의일자    -> 전표일자       날짜 기반 분석의 기준일
#  결의번호    -> 전표번호       전표 식별 키
#  사원        -> 사원명         사원별집계 분석용
#  차변금액    -> 차변           금액 분석 기준
#  대변금액    -> 대변           금액 분석 기준
#  회계일자    -> 등록일자       일자차이분석: 결의 -> 회계 지연 탐지
#  계정과목    -> 계정명         main_analyzer fallback 이 처리 (COLUMN_MAP 불필요)
#  거래처명    -> (유지)         표준명 동일
#  적요        -> (유지)         표준명 동일
#  부서/차대구분/승인요청일 등   분석 미사용, 그대로 유지
# =============================================================================
COLUMN_MAP: dict[str, str] = {
    '결의일자': '전표일자',
    '결의번호': '전표번호',
    '차변금액': '차변',
    '대변금액': '대변',
    '사원':     '사원명',
    '회계일자': '등록일자',   # 일자차이분석: 결의일자 vs 회계일자 지연 탐지
}

# PK 컬럼명 (main_analyzer 가 참조할 수 있도록 상수로 노출)
PK_COL = 'PK'


# =============================================================================
# 계정별원장 변환 (2026-11-23 blue sky 요청: 4번 TB_TO_GL_TEST 시트가 안 나오던 문제)
#
#  경남제약 ERP 원장은 시트 1개에 계정이 '계정: 11101010 : 현금' 머리행으로 이어서
#  나열되고(열: 회계일자/전표번호/적요/차변금액/대변금액/잔액), 기초잔액은 전표번호 칸이
#  '이월'인 행, 집계행은 '월계'/'누계'. 엔진은 이 형식을 모르므로 공용 더존 7열 형식
#  (계정 1개 = 시트 1개, 시트명 '순번_계정명')의 '{원본명}_변환.xlsx'로 바꾼다.
#  변환본이 있으면 엔진은 원본 .xlsx를 읽지 않는다. 원본이 안 바뀌면 재생성 안 함.
# =============================================================================
_DIR = os.path.dirname(os.path.abspath(__file__))
_LEDGER_DIRS = [os.path.join(_DIR, 'data', 'current'), os.path.join(_DIR, 'data', 'previous')]
_ACCT_HEADER_RE = re.compile(r'^계정:\s*(\S+)\s*:\s*(.*)$')
_DUZON_HEADER = ['날짜', '적요란', '코드', '거래처', '차변', '대변', '잔액']
# 엔진이 적요란에서 기초잔액·집계행을 판정하는 단어 — 실거래 적요에 들어 있으면 오판되므로 피한다
_ENGINE_KEYWORDS = ('전기이월', '기초잔액')
_ENGINE_SKIP_EXACT = ('월계', '누계', '합계')


def _ledger_sources() -> list:
    out = []
    for d in _LEDGER_DIRS:
        if os.path.isdir(d):
            out += [os.path.join(d, f) for f in sorted(os.listdir(d))
                    if f.endswith('.xlsx') and '계정별원장' in f
                    and not f.startswith('~$') and not f.endswith('_변환.xlsx')]
    return out


def _is_kyungnam_ledger(path: str) -> bool:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        rows = list(wb[wb.sheetnames[0]].iter_rows(min_row=1, max_row=2, values_only=True))
    finally:
        wb.close()
    return (len(rows) == 2 and rows[0][:2] == ('회계일자', '전표번호')
            and isinstance(rows[1][0], str) and rows[1][0].startswith('계정:'))


def _convert_ledger(src: str) -> None:
    out_path = src[:-5] + '_변환.xlsx'
    if os.path.isfile(out_path) and os.path.getmtime(out_path) >= max(os.path.getmtime(src),
                                                                     os.path.getmtime(__file__)):
        return
    if not _is_kyungnam_ledger(src):
        print(f'  [kyungnam/preprocess] 계정별원장 형식 아님 → 변환 안 함: {os.path.basename(src)}')
        return
    wb_in = openpyxl.load_workbook(src, read_only=True, data_only=True)
    wb = openpyxl.Workbook(write_only=True)
    ws, n_acct, n_open, total_dr, total_cr = None, 0, 0, 0.0, 0.0
    try:
        for r in wb_in[wb_in.sheetnames[0]].iter_rows(min_row=2, values_only=True):
            date, no, desc, dr, cr, bal = (tuple(r) + (None,) * 6)[:6]
            m = _ACCT_HEADER_RE.match(date) if isinstance(date, str) else None
            if m:
                name = m.group(2).strip()
                # 시트명 금지문자는 공백으로(엔진 계정명 매칭은 공백·/ 무시). 순번 접두어로 시트명은 유일.
                safe = re.sub(r'[\\/*?:\[\]]', ' ', name)
                ws = wb.create_sheet(f'{n_acct}_{safe}'[:31])
                ws.append(_DUZON_HEADER)
                n_acct += 1
                continue
            if ws is None or no in ('월계', '누계'):
                continue
            dr, cr = float(dr or 0), float(cr or 0)
            if no == '이월':
                ws.append([date, '전기이월', None, None, dr, cr, bal])
                n_open += 1
                continue
            text = re.sub(r'[\s\[\]]', '', str(desc or ''))
            if any(k in text for k in _ENGINE_KEYWORDS) or text in _ENGINE_SKIP_EXACT:
                desc = '거래: ' + text.replace('전기이월', '전기·이월').replace('기초잔액', '기초·잔액')
            ws.append([date, desc, no, None, dr, cr, bal])
            total_dr += dr
            total_cr += cr
    finally:
        wb_in.close()
    wb.save(out_path)
    print(f'  [kyungnam/preprocess] 계정별원장 변환: {os.path.basename(src)} → '
          f'{os.path.basename(out_path)} ({n_acct}개 계정, 이월행 {n_open}개, '
          f'거래 차변 {total_dr:,.0f} / 대변 {total_cr:,.0f})')


def convert_ledgers() -> None:
    for src in _ledger_sources():
        try:
            _convert_ledger(src)
        except Exception as e:  # 원장 변환 실패가 분개장 분석 전체를 막지 않게
            print(f'  [kyungnam/preprocess] 계정별원장 변환 실패: {os.path.basename(src)} — {e}')


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
    convert_ledgers()   # 4번 TB_TO_GL_TEST용 계정별원장 변환(원본 그대로면 건너뜀)

    print('  [kyungnam/preprocess] 컬럼 매핑 시작')

    # Step 1 — 컬럼명 표준화
    df = _map_columns(df)

    # Step 2 — 누락 표준 컬럼 보정 (분석 함수 오류 방지)
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
