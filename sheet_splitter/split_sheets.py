#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
sheet_splitter/split_sheets.py
통합 엑셀 파일을 시트별 분리/그룹핑 저장

Usage:
    # [기본] 단일 시트의 컬럼값으로 분리
    python split_sheets.py
    python split_sheets.py --mode col --col 조서번호_시트명

    # [시트명 그룹핑] '_' 앞 접두어 기준으로 시트를 묶어 출력
    python split_sheets.py --mode sheet
    python split_sheets.py --mode sheet --input 삼동산업_2025.xlsx
"""

import sys
import os
import glob
import argparse
import re

import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR  = os.path.join(BASE_DIR, 'input_data')
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')

DEFAULT_TARGET_COL   = '조서번호_시트명'
DEFAULT_NAN_LABEL    = '기타_분류안됨'
OUTPUT_FILENAME_COL  = '조서분리완료_결과물.xlsx'
OUTPUT_FILENAME_SHEET = '시트그룹핑_결과물.xlsx'
OUTPUT_FILENAME      = OUTPUT_FILENAME_COL   # 하위호환

MAX_COL_WIDTH = 60
MIN_COL_WIDTH = 8
HDR_EXTRA     = 4    # 헤더 여유 폭


# ── 파일 탐색 ─────────────────────────────────────────────────────────────────

def find_input_file(directory: str, filename: str | None = None) -> str:
    """input_data/ 에서 .xlsx 탐색. 파일명 지정 시 해당 파일, 없으면 최근 파일 자동 선택."""
    if filename:
        path = os.path.join(directory, filename)
        if os.path.isfile(path):
            return path
        raise FileNotFoundError(f'지정 파일 없음: {path}')

    candidates = sorted(
        [f for f in glob.glob(os.path.join(directory, '*.xlsx'))
         if not os.path.basename(f).startswith('~$')],
        key=os.path.getmtime,
        reverse=True,
    )
    if not candidates:
        raise FileNotFoundError(
            f'input_data/ 에 .xlsx 파일이 없습니다.\n경로: {directory}'
        )
    if len(candidates) > 1:
        print(f'[주의] 파일 {len(candidates)}개 발견 → 최근 파일 사용: '
              f'{os.path.basename(candidates[0])}')
    return candidates[0]


# ── 데이터 로드 (수식 배제, 값만 추출) ───────────────────────────────────────

def _dedup_columns(headers: list[str]) -> list[str]:
    """중복 컬럼명에 .1 .2 ... 접미사를 붙여 유일하게 만든다."""
    seen: dict[str, int] = {}
    result = []
    for h in headers:
        if h in seen:
            seen[h] += 1
            result.append(f'{h}.{seen[h]}')
        else:
            seen[h] = 0
            result.append(h)
    return result


def _read_sheet_df(filepath: str, sheet_name: str) -> pd.DataFrame:
    """단일 시트를 값만 추출해 DataFrame으로 반환 (verbose 없음)."""
    wb = load_workbook(filepath, data_only=True, read_only=True)
    if sheet_name not in wb.sheetnames:
        wb.close()
        raise ValueError(f'시트 "{sheet_name}" 없음')
    ws = wb[sheet_name]
    rows = list(ws.iter_rows(values_only=True))
    wb.close()
    if not rows:
        return pd.DataFrame()
    raw_headers = [
        str(c).strip() if c is not None else f'열_{i}'
        for i, c in enumerate(rows[0], 1)
    ]
    headers = _dedup_columns(raw_headers)
    data = [list(r) for r in rows[1:] if any(c is not None for c in r)]
    return pd.DataFrame(data, columns=headers)


def load_as_values(filepath: str, sheet_name: str | None = None) -> pd.DataFrame:
    """
    openpyxl data_only=True 로 읽어 수식·외부링크를 완전 배제하고
    계산된 결과값(캐시값)만 DataFrame으로 반환.

    sheet_name 미지정 시 활성 시트(active sheet) 사용.
    """
    print(f'  파일 로드 중: {os.path.basename(filepath)}')

    wb = load_workbook(filepath, data_only=True, read_only=True)

    if sheet_name:
        if sheet_name not in wb.sheetnames:
            raise ValueError(
                f'시트 "{sheet_name}" 없음. 사용 가능: {wb.sheetnames}'
            )
        ws = wb[sheet_name]
    else:
        ws = wb.active
        print(f'  활성 시트: {ws.title}')

    rows = list(ws.iter_rows(values_only=True))
    wb.close()

    if not rows:
        raise ValueError('시트에 데이터가 없습니다.')

    headers = [
        str(c).strip() if c is not None else f'열_{i}'
        for i, c in enumerate(rows[0], 1)
    ]
    data = [list(r) for r in rows[1:] if any(c is not None for c in r)]

    df = pd.DataFrame(data, columns=headers)
    print(f'  → {len(df):,}행 × {len(df.columns)}열 로드 완료')
    return df


# ── 시트명 정제 ───────────────────────────────────────────────────────────────

def sanitize_sheet_name(name: str) -> str:
    """엑셀 시트명 제약 처리: 금지 문자 제거, 31자 초과 시 절삭."""
    name = re.sub(r'[\\/:*?\[\]]', '_', str(name)).strip()
    return name[:31] if name else '시트'


# ── 서식 후처리 ───────────────────────────────────────────────────────────────

def apply_header_style(ws):
    """1행 헤더: 진한 파랑 배경 + 흰 볼드 + 가운데 정렬 + 틀 고정."""
    HDR_FILL = PatternFill('solid', fgColor='1F497D')
    HDR_FONT = Font(bold=True, color='FFFFFF', size=10)
    for cell in ws[1]:
        cell.fill      = HDR_FILL
        cell.font      = HDR_FONT
        cell.alignment = Alignment(horizontal='center', vertical='center',
                                   wrap_text=True)
    ws.row_dimensions[1].height = 24
    ws.freeze_panes = 'A2'


def auto_col_width(ws,
                   min_width: int = MIN_COL_WIDTH,
                   max_width: int = MAX_COL_WIDTH,
                   extra:     int = HDR_EXTRA):
    """컬럼 데이터 최대 길이 기반 열 너비 자동 조절. 한글은 2폭으로 계산."""
    for col_cells in ws.iter_cols():
        col_letter = get_column_letter(col_cells[0].column)
        max_len = 0
        for cell in col_cells:
            if cell.value is None:
                continue
            text  = str(cell.value)
            # 한글·전각문자(U+00FF 초과)는 폭 2, 그 외 1
            width = sum(2 if ord(c) > 0xFF else 1 for c in text)
            if width > max_len:
                max_len = width
        ws.column_dimensions[col_letter].width = min(
            max(max_len + extra, min_width), max_width
        )


# ── 시트명 접두어 그룹핑 (--mode sheet) ──────────────────────────────────────

NOTES_SHEET_KEYWORD = '보고서주석'
OUTPUT_FILENAME_NOTES = '보고서주석.xlsx'


def _write_groups(filepath: str, sheet_list: list[str], output_path: str, sep: str):
    """sheet_list 를 접두어로 그룹핑하여 output_path 에 저장."""
    groups: dict[str, list[str]] = {}
    for sn in sheet_list:
        prefix = sn.split(sep, 1)[0] if sep in sn else sn
        groups.setdefault(prefix, []).append(sn)

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for prefix in sorted(groups.keys()):
            sub_sheets = groups[prefix]
            dfs = []
            for sn in sub_sheets:
                try:
                    df_s = _read_sheet_df(filepath, sn)
                    if df_s.empty:
                        continue
                    df_s.insert(0, '원본시트명', sn)
                    dfs.append(df_s)
                except Exception as e:
                    print(f'  [경고] {sn}: {e}')

            if not dfs:
                continue

            combined = pd.concat(dfs, ignore_index=True, sort=False)
            out_name = sanitize_sheet_name(prefix)
            combined.to_excel(writer, sheet_name=out_name, index=False)

            ws = writer.sheets[out_name]
            apply_header_style(ws)
            auto_col_width(ws)

            sub_preview = ', '.join(sub_sheets[:3])
            if len(sub_sheets) > 3:
                sub_preview += f' 외 {len(sub_sheets)-3}개'
            print(f'  [{out_name:<31}]  {len(combined):>6,}행  ← {sub_preview}')

    return len(groups)


def group_by_sheet_prefix(
    nan_label:  str       = DEFAULT_NAN_LABEL,
    input_file: str | None = None,
    sep:        str       = '_',
):
    """
    입력 파일의 시트를 두 단계로 처리:

    1) 시트명이 정확히 '보고서주석'인 시트가 존재하면,
       그 시트부터 끝까지를 보고서주석.xlsx 로 별도 저장.
    2) 나머지(그 이전) 시트들은 '_' 앞 접두어 기준으로 그룹핑하여
       시트그룹핑_결과물.xlsx 에 저장.
    """
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    filepath = find_input_file(INPUT_DIR, input_file)
    print(f'\n  파일: {os.path.basename(filepath)}')

    wb_tmp = load_workbook(filepath, data_only=True, read_only=True)
    all_sheets = wb_tmp.sheetnames
    wb_tmp.close()

    # 정확히 '보고서주석'과 일치하는 시트 인덱스 탐색
    notes_idx = next(
        (i for i, sn in enumerate(all_sheets) if sn == NOTES_SHEET_KEYWORD),
        None,
    )

    if notes_idx is not None:
        main_sheets  = all_sheets[:notes_idx]
        notes_sheets = all_sheets[notes_idx:]
        print(f'  총 시트 수     : {len(all_sheets)}개')
        print(f'  감사조서 시트  : {len(main_sheets)}개 (접두어 그룹핑 대상)')
        print(f'  보고서주석 시트: {len(notes_sheets)}개 ({NOTES_SHEET_KEYWORD} 포함 이후)')
    else:
        main_sheets  = list(all_sheets)
        notes_sheets = []
        print(f'  총 시트 수: {len(all_sheets)}개  (보고서주석 시트 없음)')

    # ── 감사조서 부분 처리 ──────────────────────────────────────────────────
    if main_sheets:
        print(f'\n  [감사조서 그룹핑] → {OUTPUT_FILENAME_SHEET}')
        output_path = os.path.join(OUTPUT_DIR, OUTPUT_FILENAME_SHEET)
        n_groups = _write_groups(filepath, main_sheets, output_path, sep)
        print(f'  → {n_groups}개 그룹 저장 완료')

    # ── 보고서주석 부분 처리 ───────────────────────────────────────────────
    if notes_sheets:
        print(f'\n  [보고서주석 저장] → {OUTPUT_FILENAME_NOTES}')
        notes_path = os.path.join(OUTPUT_DIR, OUTPUT_FILENAME_NOTES)
        n_notes = _write_groups(filepath, notes_sheets, notes_path, sep)
        print(f'  → {n_notes}개 그룹 저장 완료')

    print('\n  전체 완료')


# ── 메인 분리 로직 (--mode col) ───────────────────────────────────────────────

def split_sheets(
    target_col: str       = DEFAULT_TARGET_COL,
    nan_label:  str       = DEFAULT_NAN_LABEL,
    input_file: str | None = None,
    sheet_name: str | None = None,
):
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # 1. 파일 탐색 및 로드
    filepath = find_input_file(INPUT_DIR, input_file)
    df = load_as_values(filepath, sheet_name)

    # 2. 분리 기준 컬럼 검증 (정확 일치 → 부분 일치 폴백)
    if target_col not in df.columns:
        close = [c for c in df.columns
                 if target_col in str(c) or str(c) in target_col]
        if close:
            print(f'[안내] 컬럼 "{target_col}" 없음 → 유사 컬럼 "{close[0]}" 사용')
            target_col = close[0]
        else:
            raise ValueError(
                f'분리 기준 컬럼 "{target_col}"을 찾을 수 없습니다.\n'
                f'사용 가능한 컬럼 목록:\n  {list(df.columns)}'
            )

    # 3. 결측치 처리 + 오름차순 정렬
    df[target_col] = df[target_col].fillna(nan_label).astype(str).str.strip()
    df.loc[df[target_col] == '', target_col] = nan_label

    unique_vals = sorted(df[target_col].unique())
    print(f'\n  분리 기준 컬럼 : {target_col}')
    print(f'  고유값 수       : {len(unique_vals)}개')
    print(f'  출력 파일       : {OUTPUT_FILENAME_COL}\n')

    output_path = os.path.join(OUTPUT_DIR, OUTPUT_FILENAME_COL)

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        for val in unique_vals:
            subset     = df[df[target_col] == val].copy()
            sheet_name = sanitize_sheet_name(val)

            # to_excel 은 순수 값만 기록 (수식 불가) → 정적 데이터 보장
            subset.to_excel(writer, sheet_name=sheet_name, index=False)

            ws = writer.sheets[sheet_name]
            apply_header_style(ws)
            auto_col_width(ws)

            print(f'  [{sheet_name:<31}]  {len(subset):>7,}행')

    print(f'\n  완료 — {len(unique_vals)}개 시트 → {output_path}')


# ── 진입점 ────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description=(
            'sheet  모드: 시트명 "_" 앞 접두어로 그룹핑 (기본값)\n'
            'col    모드: 특정 컬럼값 기준으로 시트 분리'
        ),
        formatter_class=argparse.RawTextHelpFormatter,
    )
    parser.add_argument(
        '--mode',
        choices=['sheet', 'col'],
        default='sheet',
        help='sheet: 시트명 접두어 그룹핑(기본값) | col: 컬럼값 기준 분리',
    )
    parser.add_argument(
        '--col',
        default=DEFAULT_TARGET_COL,
        help=f'[col 모드] 분리 기준 컬럼명 (기본값: "{DEFAULT_TARGET_COL}")',
    )
    parser.add_argument(
        '--nan',
        default=DEFAULT_NAN_LABEL,
        help=f'결측치 대체 레이블 (기본값: "{DEFAULT_NAN_LABEL}")',
    )
    parser.add_argument(
        '--input',
        default=None,
        help='input_data/ 내 파일명. 생략 시 가장 최근 .xlsx 자동 선택',
    )
    parser.add_argument(
        '--sheet',
        default=None,
        help='[col 모드] 읽을 시트명. 생략 시 활성 시트 사용',
    )
    parser.add_argument(
        '--sep',
        default='_',
        help='[sheet 모드] 그룹 구분자 (기본값: "_")',
    )
    args = parser.parse_args()

    if args.mode == 'sheet':
        group_by_sheet_prefix(
            nan_label=args.nan,
            input_file=args.input,
            sep=args.sep,
        )
    else:
        split_sheets(
            target_col=args.col,
            nan_label=args.nan,
            input_file=args.input,
            sheet_name=args.sheet,
        )


if __name__ == '__main__':
    main()
