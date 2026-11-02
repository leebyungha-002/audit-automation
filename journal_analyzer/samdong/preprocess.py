# -*- coding: utf-8 -*-
"""
journal_analyzer/samdong/preprocess.py
=======================================
samdong 분개장 전용 전처리.

원본 포맷 (2026-08-27 확인, 더존 계열 ERP 추정)
------------------------------------------------
컬럼: 회계단위/전표관리단위/전표기표번호/전표승인번호/행/계정과목/비용구분/
차변금액/대변금액/적요/외화차변금액/외화대변금액/통화/환율/증빙/기표자/
귀속부서/활동센터/관리항목1~8

문제
----
1. 전용 날짜 컬럼(전표일자/회계일자)이 없다. 회계일자는 관리항목 컬럼 중
   하나에 문자열로 들어있는데, 그 위치가 행(계정)마다 다르다
   (예: 여비교통비-식대 행은 관리항목2, 미지급금-카드 행은 관리항목3에
   '2025-01-01'이 위치 — 관리항목은 계정과목별로 의미가 달라지는
   가변 슬롯이라 고정 컬럼으로 신뢰할 수 없음).
2. '전표번호'에 해당하는 컬럼이 두 개다.
   - 전표기표번호: 라인(행) 단위로 유일 (예: "11-11-20250101-0002-001",
     차/대변 각 줄마다 접미어(-001,-002...)가 달라짐) → 상대계정 매칭에
     쓰면 차변/대변이 서로 다른 그룹으로 잡혀 매칭이 안 됨.
   - 전표승인번호: 같은 전표의 차/대변 라인이 공유 (예: "11-20250101-0002")
     → 상대계정분석(8번 메뉴) 등 "같은 전표에 속한 거래" 매칭에 적합
     (사용자 확인, 2026-08-27).
3. 거래처명 전용 컬럼이 없음 → 관리항목1을 거래처명으로 간주 (사용자 확인,
   2026-08-27). 단, 확보된 샘플 4행(여비교통비-식대/미지급금-카드, 비용·
   미지급금 라인)에서는 관리항목1이 '신용카드'/카드번호처럼 거래처명이
   아닌 값이었다 — 관리항목 슬롯은 계정과목별로 의미가 달라지므로, 매출
   계정(외상매출금 등)에서 관리항목1이 실제 거래처명이 맞는지는 실제
   매출 관련 데이터로 재확인 필요.
4. 은행조회서 완전성(22번 메뉴)용 금융기관명이 거래처명(관리항목1)에는 없다.
   차입금/이자비용 계정에서 관리항목1은 대출계좌번호("037-031622-32-10693"
   등)일 뿐이고, 실제 금융기관명+지점("산업 영업", "기업 남대문", "우리 천안"
   등)은 관리항목2에 들어있다 (사용자 확인, 2026-09-07). main_analyzer의
   analyze_bank_confirmation은 FI_SOURCE_COL('금융기관명_원본') 컬럼이 있으면
   거래처명 대신 그걸 우선 사용하므로, 여기서 관리항목2를 그 컬럼에 그대로
   복사해둔다 — 차입금 외 다른 계정 행의 관리항목2 값(의미가 다를 수 있음)은
   analyze_bank_confirmation이 애초에 차입금/이자비용류 계정만 필터링해서
   쓰므로 섞여도 무해하다.

해결
----
전표승인번호(예: "11-20250101-0002")를 기준으로:
  - 전표번호 / 전표그룹키 컬럼을 전표승인번호 값 그대로 생성
    (차/대변 라인이 자동으로 같은 그룹으로 묶임 → 상대계정분석 가능)
  - 문자열 중 8자리 숫자(YYYYMMDD, 예: "20250101")를 정규식으로 추출해
    전표일자 컬럼 생성. main_analyzer._preprocess_df() 가 이 문자열을
    표준 날짜 파싱 로직(YYYYMMDD 포맷)으로 이어받아 datetime으로
    변환하므로, 여기서는 파싱된 datetime이 아니라 원본 문자열째로 넘긴다
    (미리 datetime으로 바꿔두면 _preprocess_df 의 재파싱 로직이 numeric
    변환을 시도하다가 깨짐 — 다른 회사 preprocess.py들과 동일 관례).

계정별원장 변환 (2026-11-02)
----------------------------
ERP 보조원장(계정별)은 .xls 인쇄형 보고서라 main_analyzer가 읽지 못한다. 시트 1개에
페이지마다 제목·머리글이 반복되고, 금액은 콤마가 든 문자열이며, 파일이 자산/부채/
수익/제조/판관으로 나뉘어 같은 비용 계정이 제조·판관에 따로 나온다. 그래서
preprocess() 실행 시 data/current의 .xls를 계정 1개 = 시트 1개인 더존 계열 xlsx로
변환해 같은 폴더에 저장한다. 4번 데이터개요는 이 파일을 공용 로직 그대로 읽는다.
같은 계정의 여러 구간은 1개 시트로 합치고, 뒤 구간의 표시잔액에 앞 구간 최종잔액을
더해 시트 안에서 잔액이 이어지게 한다. 원본 .xls가 바뀌지 않았으면 다시 만들지 않는다.
컬럼(0-idx): 1 일자 / 3 계정명('계정과목' 머리글 행) / 8 적요 / 11 차변 / 12 대변 / 13 잔액

main_analyzer.py 호출 규약
--------------------------
함수 시그니처 :  preprocess(df: pd.DataFrame) -> pd.DataFrame
"""

import glob
import os
import re

import openpyxl
import pandas as pd
import xlrd

# main_analyzer.COL_JOURNAL_ID / COL_JOURNAL_KEY 와 동일한 값
# (순환 임포트 방지를 위해 문자열 직접 사용)
JOURNAL_ID_COL  = '전표번호'
JOURNAL_KEY_COL = '전표그룹키'
DATE_COL        = '전표일자'

_DATE_RE = re.compile(r'(20\d{6})')

_CURRENT_DIR     = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'data', 'current')
_LEDGER_OUTPUT   = '당기_samdong_계정별원장_변환.xlsx'
_LEDGER_HEADER   = ['날짜', '적요란', '코드', '거래처명', '사업자등록번호', '차변', '대변', '잔액']
_LEDGER_DATE_RE  = re.compile(r'\d{4}-\d{2}-\d{2}')
_SHEET_BAD_CHARS = re.compile(r'[\[\]:*?/\\]')
_CONTRA_KEYWORDS = ('충당금', '누계액', '국고보조금')


def _amount(value):
    """'1,234,567' -> 1234567.0, 빈 칸은 None."""
    text = str(value).replace(',', '').strip()
    return float(text) if text else None


def _read_ledger_sections(ws) -> list:
    """보조원장 시트에서 [계정명, 전월이월행, 거래행 리스트] 구간을 순서대로 뽑는다."""
    sections = []
    page_name = None
    for r in range(ws.nrows):
        row = ws.row_values(r)
        if re.sub(r'\s', '', str(row[1])) == '계정과목':
            page_name = str(row[3]).strip()
            continue
        is_open = re.sub(r'\s', '', str(row[8])) == '전월이월'
        if not is_open and not _LEDGER_DATE_RE.fullmatch(str(row[1]).strip()):
            continue                             # 제목·헤더·빈 행·페이지 번호 행
        amounts = (_amount(row[11]) or 0.0, _amount(row[12]) or 0.0, _amount(row[13]))
        if is_open:
            sections.append([page_name, amounts, []])
            continue
        if not sections or sections[-1][0] != page_name:
            sections.append([page_name, (0.0, 0.0, None), []])
        sections[-1][2].append((str(row[1]).strip(), str(row[8]).strip()) + amounts)
    return sections


def _opening_amounts(name: str, opening: tuple, txns: list) -> tuple:
    """
    전월이월 금액은 잔액 칸에만 있어(차변·대변 칸은 빈칸) 그대로 두면 공용 로직이
    기초잔액을 0으로 읽는다. 거래 행의 잔액 증감 방향으로 차변/대변 성격을 판정해
    해당 칸으로 옮긴다. 거래가 없는 계정은 차감계정 키워드로만 대변 처리한다.
    """
    debit, credit, balance = opening
    if debit or credit or not balance:
        return debit, credit, balance

    votes_debit = votes_credit = 0
    running = balance
    for _, _, t_debit, t_credit, t_balance in txns:
        if t_balance is None:
            continue
        debit_match  = abs(running + t_debit - t_credit - t_balance) <= 1
        credit_match = abs(running - t_debit + t_credit - t_balance) <= 1
        votes_debit  += debit_match and not credit_match
        votes_credit += credit_match and not debit_match
        running = t_balance

    if votes_debit == votes_credit:
        is_credit = any(k in name for k in _CONTRA_KEYWORDS)
    else:
        is_credit = votes_credit > votes_debit
    return (0.0, balance, balance) if is_credit else (balance, 0.0, balance)


def _write_ledger_workbook(sections: list, out_path: str) -> None:
    """구간들을 계정별로 합쳐 계정 1개 = 시트 1개 xlsx로 저장한다."""
    accounts = {}
    for name, opening, txns in sections:
        accounts.setdefault(name, []).append((_opening_amounts(name, opening, txns), txns))

    wb = openpyxl.Workbook(write_only=True)
    for i, (name, parts) in enumerate(accounts.items()):
        ws = wb.create_sheet(f'{i}_{_SHEET_BAD_CHARS.sub(" ", name)}'[:31])
        ws.append(_LEDGER_HEADER)
        open_debit  = sum(opening[0] for opening, _ in parts)
        open_credit = sum(opening[1] for opening, _ in parts)
        open_bal    = sum(opening[2] or 0.0 for opening, _ in parts)
        ws.append([None, '전기이월', None, None, None, open_debit, open_credit, open_bal])

        carry = 0.0                              # 앞 구간까지의 최종잔액 합계
        for opening, txns in parts:
            for date, desc, debit, credit, balance in txns:
                shown = None if balance is None else balance + carry
                ws.append([date, desc, None, None, None, debit, credit, shown])
            last = txns[-1][4] if txns else opening[2]
            carry += last or 0.0
    wb.save(out_path)


def convert_ledgers(current_dir: str = _CURRENT_DIR) -> None:
    """data/current의 .xls 보조원장을 main_analyzer가 읽는 계정별원장 xlsx로 변환한다."""
    sources = [p for p in sorted(glob.glob(os.path.join(current_dir, '*.xls')))
               if not os.path.basename(p).startswith('~$')]
    if not sources:
        return
    out_path = os.path.join(current_dir, _LEDGER_OUTPUT)
    if (os.path.exists(out_path)
            and os.path.getmtime(out_path) >= max(os.path.getmtime(p) for p in sources)):
        print(f'  [samdong/preprocess] 계정별원장 변환 생략 (원본 변경 없음): {_LEDGER_OUTPUT}')
        return

    sections = []
    for path in sources:
        with open(os.devnull, 'w') as log:       # OLE2 경고가 로그를 덮지 않도록
            ws = xlrd.open_workbook(path, logfile=log).sheet_by_index(0)
        if ws.ncols < 14 or ws.nrows < 2 or '보조원장' not in re.sub(r'\s', '', str(ws.cell_value(1, 2))):
            print(f'  [samdong/preprocess] 보조원장 형식이 아니라 건너뜀: {os.path.basename(path)}')
            continue
        sections.extend(_read_ledger_sections(ws))
    if not sections:
        return

    _write_ledger_workbook(sections, out_path)
    n_txn = sum(len(s[2]) for s in sections)
    sum_debit  = sum(t[2] for s in sections for t in s[2])
    sum_credit = sum(t[3] for s in sections for t in s[2])
    print(f'  [samdong/preprocess] 계정별원장 변환 완료: {len(sources)}개 파일, '
          f'{len({s[0] for s in sections})}개 계정, 거래 {n_txn}행, '
          f'차변 {sum_debit:,.0f} / 대변 {sum_credit:,.0f}')


def preprocess(df: pd.DataFrame) -> pd.DataFrame:
    convert_ledgers()

    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    if '전표승인번호' not in df.columns:
        print('  [samdong/preprocess] 전표승인번호 컬럼 없음 → 전처리 건너뜀')
        return df

    approval = df['전표승인번호'].astype(str).str.strip()

    # 전표승인번호를 전표번호/전표그룹키로 그대로 사용
    # (같은 전표의 차/대변 라인이 이 값을 공유하므로 상대계정분석에 바로 적합)
    df[JOURNAL_ID_COL]  = approval
    df[JOURNAL_KEY_COL] = approval

    # 전표승인번호 안의 YYYYMMDD 추출 → 전표일자 (문자열째로 넘겨 표준 파싱에 위임)
    date_str = approval.str.extract(_DATE_RE, expand=False)
    df[DATE_COL] = date_str

    # 거래처명 전용 컬럼이 없는 회사 → 관리항목1을 거래처명으로 간주 (사용자 확인, 2026-08-27)
    if '거래처명' not in df.columns:
        if '관리항목1' in df.columns:
            df['거래처명'] = df['관리항목1']
        else:
            df['거래처명'] = ''
    # 기표자 → 사원명 (사원별집계 메뉴 인식용)
    if '사원명' not in df.columns and '기표자' in df.columns:
        df = df.rename(columns={'기표자': '사원명'})

    # 은행조회서완전성(22번)용 금융기관명 원본 — 관리항목2에 있음 (사용자 확인, 2026-09-07)
    if '관리항목2' in df.columns:
        df['금융기관명_원본'] = df['관리항목2']

    n_ok = date_str.notna().sum()
    print(f'  [samdong/preprocess] 전표승인번호 기반 전표번호/전표그룹키/전표일자 생성 완료 '
          f'(날짜 추출 성공 {n_ok}/{len(df)}행)')
    return df
