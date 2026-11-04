# -*- coding: utf-8 -*-
"""
journal_analyzer/samdong_material/convert_prev_detail.py
========================================================
samdong_material 전기 조서용 명세서(2026-11-04 업로드) → 엔진용 전기 계정별_거래처별 명세 변환.

입력: data/previous/명세서원본/ 의 조서 양식 명세서
  (02) 2025년 당좌 자산-삼동소재.xlsx   — 매출채권 시트(외상매출금·받을어음)
  (08) 2025년 유동 부채-삼동소재.xlsx   — 매입채무(외상매입금)·단기차입금 시트
  (09) 2025년 비유동 부채-삼동소재.xlsx — 장기차입금 시트
  (원본을 data/previous 바로 아래 두면 load_data가 전기 분개장으로 읽으므로 하위 폴더에 둔다)

출력: data/previous/전기_계정별_거래처별_명세_samdong_material.xlsx
  계정별 시트, 열 = 거래처명 | 잔액 | 비고 (20번 잔액증감분석 대상 5개 계정만)

거래처명은 당기 분개장 표기에 맞춘다(㈜→(주), 기업은행 남대문→기업은행 남대문지점,
중진공→중소기업진흥공단). 단기차입금 #00033(10억)은 명세서에 ㈜삼동산업으로 적혀 있으나
기업은행 차입금이 맞다(blue sky 확인, 2026-11-04).

검증: 계정별 합계를 명세서 총계 행, 전기 계정별원장 기말잔액과 대사한다.
실행: python convert_prev_detail.py
"""
import os
import re

import openpyxl
import pandas as pd

_DIR = os.path.dirname(os.path.abspath(__file__))
_PREV_DIR = os.path.join(_DIR, 'data', 'previous')
_SRC_DIR = os.path.join(_PREV_DIR, '명세서원본')
_OUT = os.path.join(_PREV_DIR, '전기_계정별_거래처별_명세_samdong_material.xlsx')
_PREV_LEDGER = os.path.join(_PREV_DIR, '전기_samdong_material_계정별원장_25년.xlsx')

# (출력 계정명, 원본 파일 접두, 시트명, 계정 열, 거래처 열, 금액 열, 비고 열)
# 계정 열이 있으면 그 값(앞 행 이어받기)으로 계정을 구분, None이면 시트 전체가 한 계정.
_SPECS = [
    ('외상매출금', '(02)', '매출채권', 'A', 'C', 'G', None),
    ('받을어음', '(02)', '매출채권', 'A', 'C', 'G', None),
    ('외상매입금', '(08)', '매입채무', 'A', 'C', 'G', None),
    ('단기차입금', '(08)', '단기차입금', None, 'E', 'G', 'B'),
    ('장기차입금', '(09)', '장기차입금', None, 'F', 'I', 'B'),
]

_NAME_MAP = {
    '기업은행 남대문': '기업은행 남대문지점',
    '중진공': '중소기업진흥공단',
}
# 관리번호별 차입처 정정(명세서 기재 오류)
_LENDER_OVERRIDE = {'#00033': '기업은행 남대문지점'}


def _clean_name(v) -> str:
    s = re.sub(r'\s+', ' ', str(v)).strip().replace('㈜', '(주)')
    return _NAME_MAP.get(s, s)


def _is_total_row(ws, r) -> bool:
    for col in 'AB':
        v = ws[f'{col}{r}'].value
        if isinstance(v, str) and re.sub(r'\s+', '', v) in ('소계', '총계', '합계'):
            return True
    return False


def _find_src(prefix):
    for f in sorted(os.listdir(_SRC_DIR)):
        if f.startswith(prefix) and f.endswith('.xlsx') and not f.startswith('~$'):
            return os.path.join(_SRC_DIR, f)
    raise FileNotFoundError(f'{_SRC_DIR} 에 {prefix} 명세서 없음')


def _extract(spec):
    acct, prefix, sheet, acct_col, vend_col, amt_col, note_col = spec
    ws = openpyxl.load_workbook(_find_src(prefix), data_only=True)[sheet]
    rows, cur_acct, total = [], None, None
    for r in range(1, ws.max_row + 1):
        a = ws[f'A{r}'].value
        if isinstance(a, str) and re.sub(r'\s+', '', a) == '총계':
            total = ws[f'{amt_col}{r}'].value
            break
        if acct_col and isinstance(a, str) and a.strip() and not _is_total_row(ws, r):
            cur_acct = re.sub(r'\s+', '', a)
        if _is_total_row(ws, r):
            continue
        amt = ws[f'{amt_col}{r}'].value
        if not isinstance(amt, (int, float)):
            continue
        if acct_col and cur_acct != acct:
            continue
        note = str(ws[f'{note_col}{r}'].value).strip() if note_col and ws[f'{note_col}{r}'].value else ''
        vend = _LENDER_OVERRIDE.get(note) or _clean_name(ws[f'{vend_col}{r}'].value or '')
        rows.append({'거래처명': vend, '잔액': float(amt), '비고': note})
    # 매출채권 시트처럼 한 시트에 여러 계정이 있으면 총계는 계정 합계가 아니므로 비교 생략
    return pd.DataFrame(rows, columns=['거래처명', '잔액', '비고']), (None if acct_col else total)


def _ledger_closing(acct):
    """전기 계정별원장(7열 더존) 해당 계정 시트의 마지막 잔액."""
    if not os.path.isfile(_PREV_LEDGER):
        return None
    wb = openpyxl.load_workbook(_PREV_LEDGER, read_only=True, data_only=True)
    try:
        for sn in wb.sheetnames:
            m = re.match(r'^\d+_(.+)\(\d+\)$', sn.strip())
            if m and m.group(1).strip() == acct:
                last = None
                for row in wb[sn].iter_rows(min_row=2, values_only=True):
                    if len(row) >= 7 and isinstance(row[6], (int, float)):
                        last = row[6]
                return last
    finally:
        wb.close()
    return None


def main():
    results, ok = {}, True
    for spec in _SPECS:
        acct = spec[0]
        df, total = _extract(spec)
        results[acct] = df
        s = df['잔액'].sum()
        led = _ledger_closing(acct)
        chk = []
        if total is not None:
            chk.append(f'명세 총계 {"일치" if abs(s - total) < 1 else f"불일치({total:,.0f})"}')
        if led is not None:
            chk.append(f'전기원장 {"일치" if abs(s - led) < 1 else f"불일치({led:,.0f})"}')
        else:
            chk.append('전기원장 계정 없음')
        if any('불일치' in c for c in chk):
            ok = False
        print(f'{acct}: {len(df)}행 합계 {s:,.0f} — {", ".join(chk)}')

    with pd.ExcelWriter(_OUT, engine='openpyxl') as w:
        for acct, df in results.items():
            df.to_excel(w, sheet_name=acct, index=False)
    print(f'저장: {_OUT}')
    if not ok:
        print('⚠️ 대사 불일치 계정 있음 — 확인 필요')


if __name__ == '__main__':
    main()
