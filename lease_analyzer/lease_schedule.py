#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
K-IFRS 1116 리스 회계처리 자동화 스크립트

Usage:
    python lease_schedule.py

입력: input_data/lease_{회사명}_information_{연도}.xlsx  (header=1)
출력: output/lease_schedule_{회사명}_{연도}.xlsx
  - Sheet 1: 계약별 요약
  - Sheet 2: 월별 상각 스케줄

주요 컬럼:
    리스계약번호, 리스개시일, 리스종료일, 최종 산정리스기간(월수),
    정기리스료, 지급주기(월/분기/년), 지급시점(기초/기말),
    적용할인율(연 이자율), 보증잔존가치/매수선택권,
    선급리스료, 수취한 리스인센티브, 리스개설직접원가, 복구충당부채
"""

import glob
import os
import re
import sys

import numpy_financial as npf
import pandas as pd
from dateutil.relativedelta import relativedelta
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR  = os.path.join(BASE_DIR, 'input_data')
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')

MONEY_FMT  = '#,##0'
RATE_FMT   = '0.00%'


# ── 이자율 변환 ───────────────────────────────────────────────────────────────

def _months_per_period(freq: str) -> int:
    freq = str(freq).strip()
    if freq in ('월', '월별', 'M', '매월'):                return 1
    if freq in ('분기', '분기별', 'Q', '매분기'):           return 3
    if freq in ('반기', '반기별', 'H', '매반기', '6개월'):  return 6
    if freq in ('년', '연', '연간', '년간', 'Y', '매년'):   return 12
    raise ValueError(f'지원하지 않는 지급주기: {freq}')


def _period_rate(annual_rate: float, freq: str) -> float:
    """연 이자율 → 지급주기 이자율 (유효이자율 방식)"""
    m = _months_per_period(freq)
    return (1 + annual_rate) ** (m / 12) - 1


def _monthly_rate(annual_rate: float) -> float:
    return (1 + annual_rate) ** (1 / 12) - 1


def _safe_float(v, default: float = 0.0) -> float:
    """NaN/None/공백 → default"""
    try:
        f = float(v)
        return default if (f != f) else f   # NaN check
    except (TypeError, ValueError):
        return default


def _col(row: dict, *candidates) -> object:
    """컬럼명 후보 중 값이 있는 첫 번째 반환 (괄호 제거 부분 일치 폴백)"""
    for key in candidates:
        if key in row:
            v = row[key]
            if v is not None and str(v).strip() not in ('', 'nan', 'NaN', 'None'):
                return v
    # 폴백: 괄호·특수문자 제거 후 부분 일치
    def _strip(s):
        return re.sub(r'[\(\)（）/\s]', '', str(s))
    for key in candidates:
        k_stripped = _strip(key)
        for col in row:
            if k_stripped in _strip(col) or _strip(col) in k_stripped:
                v = row[col]
                if v is not None and str(v).strip() not in ('', 'nan', 'NaN', 'None'):
                    return v
    return None


# ── 스케줄 생성 ───────────────────────────────────────────────────────────────

def build_schedule(c: dict) -> tuple:
    """
    단일 계약의 월별 상각 스케줄 생성.
    Returns (schedule_df, initial_liability, initial_rou)
    """
    start    = pd.Timestamp(c['리스개시일']).date()
    n_months = int(_safe_float(_col(c, '최종 산정리스기간(월수)'), 1))
    payment  = _safe_float(_col(c, '정기리스료'))
    freq     = str(_col(c, '지급주기(월/분기/년)', '지급주기(월/분기/반기/년)', '지급주기') or '월').strip()
    timing   = str(_col(c, '지급시점(기초/기말)', '지급시점') or '기말').strip()
    annual_r = _safe_float(_col(c, '적용할인율(연 이자율)', '적용할인율', '이자율'))
    if annual_r > 1:        # % 단위 입력 (예: 5.0 → 0.05)
        annual_r /= 100

    residual    = _safe_float(_col(c, '보증잔존가치/매수선택권', '보증잔존가치/매수선택권금액', '잔존가치'))
    prepaid     = _safe_float(_col(c, '선급리스료'))
    incentive   = _safe_float(_col(c, '수취한 리스인센티브', '리스인센티브'))
    direct_cost = _safe_float(_col(c, '리스개설직접원가', '직접원가'))
    restoration = _safe_float(_col(c, '복구충당부채', '복구충당부채(현재가치)', '복구충당'))

    mpp    = _months_per_period(freq)      # 지급주기(개월)
    n_prd  = n_months // mpp              # 총 지급 횟수
    r_prd  = _period_rate(annual_r, freq) # 기간 이자율
    r_mo   = _monthly_rate(annual_r)      # 월 이자율
    when   = 1 if timing == '기초' else 0  # numpy_financial 규칙

    # ── 최초 리스부채 현재가치 (numpy_financial.pv)
    # npf.pv: pmt 음수(유출) 입력 시 양수(자산/부채 가치)를 반환
    initial_liability = float(npf.pv(
        rate=r_prd, nper=n_prd, pmt=-payment, fv=-residual, when=when
    ))

    # ── 사용권자산 최초 원가
    initial_rou = max(0.0, initial_liability + prepaid - incentive + direct_cost + restoration)

    # ── 지급 월 집합 (1-based 월 인덱스)
    pay_months: set = set()
    if when == 1:   # 기초: 1, mpp+1, 2*mpp+1, ...
        for i in range(n_prd):
            pay_months.add(i * mpp + 1)
    else:           # 기말: mpp, 2*mpp, ...
        for i in range(1, n_prd + 1):
            pay_months.add(i * mpp)

    monthly_dep = initial_rou / n_months if n_months > 0 else 0.0   # 정액법 월 감가상각비

    rows = []
    liab = initial_liability
    rou  = initial_rou

    for mi in range(1, n_months + 1):
        dt       = start + relativedelta(months=mi - 1)
        ym       = dt.strftime('%Y-%m')
        is_pay   = mi in pay_months
        is_last  = mi == n_months

        op_liab = liab
        op_rou  = rou

        # 현금지급액
        cash = payment if is_pay else 0.0
        if is_last and residual > 0:
            cash += residual

        # 이자비용
        # 기초 지급 월: 지급 후 잔액에 이자 적용 (당월 지급액은 원금에서 선차감)
        if when == 1 and is_pay:
            interest = max(0.0, op_liab - payment) * r_mo
        else:
            interest = op_liab * r_mo

        # 기말 부채 잔액: 기초 + 이자 - 지급액 (음수 방지)
        cl_liab = max(0.0, op_liab + interest - cash)
        cl_rou  = max(0.0, op_rou - monthly_dep)

        rows.append({
            '연월':               ym,
            '기초부채잔액':        op_liab,
            '이자비용':           interest,
            '현금지급액':         cash,
            '원금상환액':         cash - interest,   # 음수=이자만 발생(비지급월)
            '기말부채잔액':        cl_liab,
            '기초자산잔액':        op_rou,
            '감가상각비':         monthly_dep,
            '기말자산잔액':        cl_rou,
            '유동성대체대상액':     None,
            '비유동성리스부채잔액':  None,
        })

        liab = cl_liab
        rou  = cl_rou

    df = pd.DataFrame(rows)

    # ── 12월 말 유동성 대체 계산
    # 당해 12월 잔액 - 다음 해 12월 잔액 = 내년도 원금상환 예정액(유동성 대체액)
    dec_idx = df[df['연월'].str.endswith('-12')].index.tolist()
    for k, idx in enumerate(dec_idx):
        cur_bal  = df.loc[idx, '기말부채잔액']
        next_bal = df.loc[dec_idx[k + 1], '기말부채잔액'] if k + 1 < len(dec_idx) else 0.0
        df.loc[idx, '유동성대체대상액']    = cur_bal - next_bal
        df.loc[idx, '비유동성리스부채잔액'] = next_bal

    return df, initial_liability, initial_rou


# ── Excel 저장 ────────────────────────────────────────────────────────────────

_HDR_FILL  = PatternFill('solid', fgColor='1F497D')
_HDR_FONT  = Font(bold=True, color='FFFFFF', size=10)
_DEC_FILL  = PatternFill('solid', fgColor='FFF2CC')   # 12월 행 강조
_CUR_FILL  = PatternFill('solid', fgColor='E2EFDA')   # 유동성 값 있는 셀


def _set_header(ws):
    for cell in ws[1]:
        cell.fill      = _HDR_FILL
        cell.font      = _HDR_FONT
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    ws.row_dimensions[1].height = 30


def _money_cols(ws, df: pd.DataFrame, exclude=(), min_row=2):
    """금액 컬럼에 천단위 콤마 서식 적용"""
    for ci, col in enumerate(df.columns, 1):
        if col in exclude:
            continue
        for cell in ws.iter_rows(min_row=min_row, max_row=ws.max_row,
                                  min_col=ci, max_col=ci):
            for c in cell:
                if isinstance(c.value, (int, float)):
                    c.number_format = MONEY_FMT


def _col_width(ws, widths: dict, default=14):
    for ci, col_dim in ws.column_dimensions.items():
        pass  # just initialize
    for ci, col in enumerate(ws.iter_cols(min_row=1, max_row=1), 1):
        hdr = str(col[0].value or '')
        ws.column_dimensions[get_column_letter(ci)].width = widths.get(hdr, default)


def save_results(summary_df: pd.DataFrame, detail_df: pd.DataFrame, output_path: str):
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:

        # ── Sheet 1: 계약별 요약 ─────────────────────────────────────────────
        summary_df.to_excel(writer, sheet_name='계약별 요약', index=False)
        ws1 = writer.sheets['계약별 요약']
        _set_header(ws1)
        _money_cols(ws1, summary_df, exclude=('리스계약번호', '리스개시일', '리스종료일',
                                               '리스기간(월)', '지급주기', '지급시점', '할인율(연)'))

        # 할인율 % 서식
        rate_ci = summary_df.columns.get_loc('할인율(연)') + 1 if '할인율(연)' in summary_df.columns else None
        if rate_ci:
            for cell in ws1.iter_rows(min_row=2, max_row=ws1.max_row,
                                       min_col=rate_ci, max_col=rate_ci):
                for c in cell:
                    c.number_format = RATE_FMT

        s_widths = {
            '리스계약번호': 18, '리스개시일': 12, '리스종료일': 12,
            '리스기간(월)': 10, '정기리스료': 14, '지급주기': 8, '지급시점': 8, '할인율(연)': 10,
            '최초 리스부채': 16, '사용권자산(최초)': 16,
        }
        _col_width(ws1, s_widths, default=18)

        # ── Sheet 2: 월별 상각 스케줄 ────────────────────────────────────────
        detail_df.to_excel(writer, sheet_name='월별 상각 스케줄', index=False)
        ws2 = writer.sheets['월별 상각 스케줄']
        _set_header(ws2)
        _money_cols(ws2, detail_df, exclude=('리스계약번호', '연월'))

        # 12월 행 배경색 + 유동성 셀 강조
        ym_ci   = detail_df.columns.get_loc('연월') + 1
        cur_ci  = detail_df.columns.get_loc('유동성대체대상액') + 1    if '유동성대체대상액'    in detail_df.columns else None
        ncur_ci = detail_df.columns.get_loc('비유동성리스부채잔액') + 1 if '비유동성리스부채잔액' in detail_df.columns else None
        n_cols  = len(detail_df.columns)

        for row_idx, ym_val in enumerate(detail_df['연월'], start=2):
            if str(ym_val).endswith('-12'):
                for ci in range(1, n_cols + 1):
                    ws2.cell(row=row_idx, column=ci).fill = _DEC_FILL
                if cur_ci:
                    ws2.cell(row=row_idx, column=cur_ci).fill  = _CUR_FILL
                    ws2.cell(row=row_idx, column=cur_ci).font  = Font(bold=True)
                if ncur_ci:
                    ws2.cell(row=row_idx, column=ncur_ci).fill = _CUR_FILL
                    ws2.cell(row=row_idx, column=ncur_ci).font = Font(bold=True)

        d_widths = {
            '리스계약번호': 18, '연월': 10,
            '기초부채잔액': 16, '이자비용': 14, '현금지급액': 14,
            '원금상환액': 14, '기말부채잔액': 16,
            '기초자산잔액': 16, '감가상각비': 14, '기말자산잔액': 16,
            '유동성대체대상액': 18, '비유동성리스부채잔액': 20,
        }
        _col_width(ws2, d_widths, default=14)

        # 행 고정 (헤더 + 계약번호/연월)
        ws2.freeze_panes = 'C2'

    print(f'  → {os.path.basename(output_path)} 저장 완료')


# ── 메인 ─────────────────────────────────────────────────────────────────────

def main():
    pattern = os.path.join(INPUT_DIR, 'lease_*_information_*.xlsx')
    files   = sorted(f for f in glob.glob(pattern)
                     if not os.path.basename(f).startswith('~$'))

    if not files:
        print(f'[안내] 처리할 파일 없음 — 경로: {INPUT_DIR}')
        print(f'       파일명 형식: lease_{{회사명}}_information_{{연도}}.xlsx')
        return

    for fpath in files:
        fname = os.path.basename(fpath)
        # 연도: 4자리 숫자(2025) 또는 fy##(fy25 → 2025) 모두 허용
        m = re.match(r'^lease_(.+)_information_(\d{4}|fy\d{2})\.xlsx$', fname, re.IGNORECASE)
        if not m:
            print(f'[건너뜀] 파일명 형식 불일치: {fname}')
            continue

        company  = m.group(1)
        raw_year = m.group(2)
        year     = raw_year if raw_year.isdigit() else f'20{raw_year[2:]}'  # fy25 → 2025
        print(f'\n{"="*60}')
        print(f'  처리: {fname}  (회사: {company}, 연도: {year})')
        print(f'{"="*60}')

        try:
            df_in = pd.read_excel(fpath, header=1)
        except Exception as e:
            print(f'  [오류] 파일 읽기 실패: {e}')
            continue

        df_in = df_in.dropna(subset=['리스계약번호'])
        df_in = df_in[df_in['리스계약번호'].astype(str).str.strip() != '']
        print(f'  계약 수: {len(df_in)}건')

        all_schedules = []
        summary_rows  = []

        for _, row in df_in.iterrows():
            cid = str(row['리스계약번호']).strip()
            print(f'  ▷ {cid}', end='  ')
            try:
                sched, init_liab, init_rou = build_schedule(row.to_dict())
            except Exception as e:
                print(f'[오류] {e}')
                continue

            sched.insert(0, '리스계약번호', cid)
            all_schedules.append(sched)

            # 당해연도 집계
            yr_rows   = sched[sched['연월'].str.startswith(year)]
            dec_row   = sched[sched['연월'] == f'{year}-12']
            has_dec   = not dec_row.empty

            yr_int    = yr_rows['이자비용'].sum()
            yr_dep    = yr_rows['감가상각비'].sum()
            yr_cash   = yr_rows['현금지급액'].sum()
            yr_end_lb = dec_row['기말부채잔액'].values[0]  if has_dec else None
            cur_p     = dec_row['유동성대체대상액'].values[0]    if has_dec else None
            nc_p      = dec_row['비유동성리스부채잔액'].values[0] if has_dec else None

            row_d    = row.to_dict()
            annual_r = _safe_float(_col(row_d, '적용할인율(연 이자율)', '적용할인율', '이자율'))
            if annual_r > 1:
                annual_r /= 100

            summary_rows.append({
                '리스계약번호':          cid,
                '리스개시일':           pd.Timestamp(row['리스개시일']).date(),
                '리스종료일':           pd.Timestamp(row['리스종료일']).date(),
                '리스기간(월)':         int(_safe_float(_col(row_d, '최종 산정리스기간(월수)'))),
                '정기리스료':           _safe_float(_col(row_d, '정기리스료')),
                '지급주기':             str(_col(row_d, '지급주기(월/분기/년)', '지급주기(월/분기/반기/년)', '지급주기') or '').strip(),
                '지급시점':             str(_col(row_d, '지급시점(기초/기말)', '지급시점') or '').strip(),
                '할인율(연)':           annual_r,
                '최초 리스부채':        round(init_liab),
                '사용권자산(최초)':     round(init_rou),
                f'{year}년 이자비용':   round(yr_int),
                f'{year}년 감가상각비': round(yr_dep),
                f'{year}년 리스료지급': round(yr_cash),
                # 만료 계약은 0으로 표시 (NaN 대신)
                f'{year}년말 리스부채': round(yr_end_lb) if yr_end_lb is not None else 0,
                '유동성대체대상액':      round(cur_p) if cur_p  is not None else 0,
                '비유동성리스부채잔액':  round(nc_p)  if nc_p   is not None else 0,
            })

            print(f'초기부채 {init_liab:>14,.0f}원 / 사용권자산 {init_rou:>14,.0f}원')

        if not all_schedules:
            print('  [안내] 처리된 계약 없음')
            continue

        summary_df = pd.DataFrame(summary_rows)

        # 전체 월별 스케줄 병합 + 금액 반올림
        detail_df = pd.concat(all_schedules, ignore_index=True)
        for col in detail_df.columns:
            if col not in ('리스계약번호', '연월'):
                detail_df[col] = pd.to_numeric(detail_df[col], errors='coerce').round(0)

        output_path = os.path.join(OUTPUT_DIR, f'lease_schedule_{company}_{year}.xlsx')
        save_results(summary_df, detail_df, output_path)

        print(f'\n  완료 — 계약 {len(summary_rows)}건 / 스케줄 {len(detail_df):,}행')
        print('=' * 60)


if __name__ == '__main__':
    main()
