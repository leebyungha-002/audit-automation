#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
jymts 이자비용 적정성 분석 입력 파일 생성 (2026-11-16)

실행: python account_analyzer/interest_analyzer/jymts/build_input.py
      → python account_analyzer/interest_analyzer/interest_expense_analysis.py jymts

원천: journal_analyzer/jymts/data/current/의 계정별원장 변환본(convert_journal.py가 만든 7열 양식)
      + 당기 분개장 변환본(기말일 = 마지막 전표일자)

jymts 원장은 계정 단위(단기차입금/장기차입금/유동성장기부채)라 차입금별로 안 나뉘어 있다.
대출은 거래처명으로 구분된다 → 거래처의 대출번호('32-00031' 형식, 없으면 거래처명 그대로)를
키로 세 계정의 거래를 모아 차입금 1개 = 시트 1개로 만든다. 같은 대출의 유동성대체
(장기차입금 대변 +X / 유동성장기부채 대변 -X)는 한 시트 안에서 상계된다.
이자만 내고 당기 원금 변동이 없는 대출(이자비용 원장에만 나옴)도 빈 시트로 만든다.

원장 '전기이월'은 계정 합계뿐이라 차입금별 기초잔액은 알 수 없다 →
  이자율설정.xlsx '이자율' 시트의 기초잔액·연이자율은 감사인이 차입금 명세로 입력.
  '설정' 시트의 원장기초잔액합계(세 계정 전기이월 합계)로 입력값 합계를 분석 시 대사.

출력 (interest_analyzer/jymts/input/):
  차입금현황.xlsx — 매번 다시 만듦(원장에서 파생)
  이자율설정.xlsx     — 없을 때만 만듦(감사인 입력값 보존). 새 차입금이 생기면 행 추가 안내만 출력
"""
import os
import re
import sys

import openpyxl
import pandas as pd

_DIR = os.path.dirname(os.path.abspath(__file__))
_REPO = os.path.abspath(os.path.join(_DIR, '..', '..', '..'))
_JOURNAL_DATA = os.path.join(_REPO, 'journal_analyzer', 'jymts', 'data', 'current')
_INPUT_DIR = os.path.join(_DIR, 'input')
_LOAN_ACCOUNTS = ('단기차입금', '장기차입금', '유동성장기부채')
_INTEREST_ACCOUNT = '이자비용'
_LOAN_NO_RE = re.compile(r'\d{2}-\d{5}')
_COLS = ['계정과목', '날짜', '적요란', '코드', '거래처', '차변', '대변']


def _find(must: tuple) -> str:
    for f in sorted(os.listdir(_JOURNAL_DATA)):
        if not f.startswith('~$') and f.endswith('_변환.xlsx') and all(m in f for m in must):
            return os.path.join(_JOURNAL_DATA, f)
    raise FileNotFoundError(f'{_JOURNAL_DATA}에 {must} 변환본 없음 — journal_analyzer로 jymts를 먼저 실행(자동 변환)')


def _account(sheet_title: str) -> str:
    """'39_단기차입금(0260)' → '단기차입금'"""
    return re.sub(r'\(\d+\)$', '', re.sub(r'^\d+_', '', sheet_title)).strip()


def _loan_key(client: str) -> str:
    m = _LOAN_NO_RE.search(client)
    return m.group(0) if m else client.strip()


def _sheet_name(client: str, used: set) -> str:
    name = re.sub(r'[\\/*?:\[\]]', ' ', client).strip()[:31]
    base, i = name, 2
    while name in used or name == _INTEREST_ACCOUNT:
        name = f'{base[:28]}_{i}'
        i += 1
    used.add(name)
    return name


def main() -> int:
    ledger = _find(('계정별원장',))
    journal = _find(('분개장',))
    jdates = pd.to_datetime(pd.read_excel(journal, usecols=['전표일자'])['전표일자'], errors='coerce')
    period_end = jdates.max().normalize()
    year = period_end.year
    period_start = pd.Timestamp(year=year, month=1, day=1)
    print(f'원장: {os.path.basename(ledger)}\n기간: {period_start:%Y-%m-%d} ~ {period_end:%Y-%m-%d} (분개장 마지막 전표일)')

    wb = openpyxl.load_workbook(ledger, read_only=True, data_only=True)
    loans, interest, ledger_open = {}, [], {}   # loans: key → {'name', 'rows', 'accounts'}
    for ws in wb.worksheets:
        acct = _account(ws.title)
        if acct not in _LOAN_ACCOUNTS and acct != _INTEREST_ACCOUNT:
            continue
        for r in ws.iter_rows(min_row=2, values_only=True):
            date, desc, code, client, dr, cr = r[0], r[1], r[2], r[3], r[4] or 0, r[5] or 0
            label = re.sub(r'[\s\[\]]', '', str(desc or ''))
            if date is None:   # 전기이월/월계/누계
                if '전기이월' in label and acct in _LOAN_ACCOUNTS:
                    ledger_open[acct] = ledger_open.get(acct, 0) + cr - dr
                continue
            d = pd.Timestamp(f'{year}-{date}')
            row = [acct, d, desc, code, client, dr or None, cr or None]
            if acct == _INTEREST_ACCOUNT:
                interest.append(row)
                if not client:
                    continue
            elif not client:
                print(f'  [경고] {acct} {date} 거래처 없는 차입금 거래 — 차입금 배정 불가: {desc} {dr} {cr}')
                continue
            # 이자 전표는 적요의 대출번호가 우선 — 거래처를 잘못 고른 전표가 있음
            # (예: 01-05 적요 '이자-32-00063'인데 거래처 '기업대출(32-00066)')
            m = _LOAN_NO_RE.search(str(desc or '')) if acct == _INTEREST_ACCOUNT else None
            key = m.group(0) if m else _loan_key(client)
            if m and _loan_key(client) != key:
                print(f'  [알림] 이자비용 {date} 적요 {key} ≠ 거래처 {client} — 적요 기준으로 배정: {dr - cr:,.0f}')
            loan = loans.setdefault(key, {'name': client, 'rows': [], 'accounts': set(), 'interest': 0})
            if _loan_key(loan['name']) != key and _loan_key(client) == key:
                loan['name'] = client   # 이름은 대출번호가 맞는 거래처로
            if acct == _INTEREST_ACCOUNT:
                loan['interest'] += dr - cr
            else:
                loan['rows'].append(row)
                loan['accounts'].add(acct)
    wb.close()

    os.makedirs(_INPUT_DIR, exist_ok=True)
    used, sheets = set(), {}
    for key, ln in sorted(loans.items(), key=lambda kv: kv[1]['name']):
        ln['sheet'] = _sheet_name(ln['name'], used)
        sheets[ln['sheet']] = pd.DataFrame(ln['rows'], columns=_COLS)
    out = os.path.join(_INPUT_DIR, '차입금현황.xlsx')
    with pd.ExcelWriter(out, engine='openpyxl') as w:
        for name, df in sheets.items():
            df.to_excel(w, sheet_name=name, index=False)
        pd.DataFrame(interest, columns=_COLS).to_excel(w, sheet_name=_INTEREST_ACCOUNT, index=False)
    total_open = sum(ledger_open.values())
    print(f'저장: {out}\n  차입금 {len(sheets)}개, 이자비용 {len(interest)}줄 '
          f'(합계 {sum((r[5] or 0) - (r[6] or 0) for r in interest):,.0f})')
    print(f'  원장 차입금 기초잔액: ' + ', '.join(f'{k} {v:,.0f}' for k, v in ledger_open.items())
          + f' → 합계 {total_open:,.0f}')

    settings = os.path.join(_INPUT_DIR, '이자율설정.xlsx')
    if os.path.exists(settings):
        have = set(pd.read_excel(settings, sheet_name='이자율')['차입금명'].astype(str).str.strip())
        new = [ln['sheet'] for ln in loans.values() if ln['sheet'] not in have]
        print(f'이자율설정.xlsx 있음 — 감사인 입력값 보존(덮어쓰지 않음)'
              + (f'. 새 차입금 행 추가 필요: {new}' if new else ''))
        return 0
    rate_rows = [{
        '차입금명': ln['sheet'], '연이자율': None, '기초잔액': None,
        '비고': (f"원장 계정: {', '.join(sorted(ln['accounts'])) or '(당기 원금 변동 없음)'} / "
               f"당기 이자비용 {ln['interest']:,.0f}"),
    } for ln in sorted(loans.values(), key=lambda x: x['sheet'])]
    with pd.ExcelWriter(settings, engine='openpyxl') as w:
        pd.DataFrame(rate_rows).to_excel(w, sheet_name='이자율', index=False)
        pd.DataFrame([
            {'항목': '기초일', '값': period_start.to_pydatetime(), '비고': None},
            {'항목': '기말일', '값': period_end.to_pydatetime(), '비고': '당기 분개장 마지막 전표일'},
            {'항목': '원장기초잔액합계', '값': total_open,
             '비고': '원장 전기이월 합계(' + ', '.join(f'{k} {v:,.0f}' for k, v in ledger_open.items())
                   + ') — 이자율 시트 기초잔액 합계와 대사'},
        ]).to_excel(w, sheet_name='설정', index=False)
    print(f'저장: {settings} — 연이자율·기초잔액을 차입금 명세 보고 입력하세요')
    return 0


if __name__ == '__main__':
    sys.exit(main())
