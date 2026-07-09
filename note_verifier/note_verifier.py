#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
note_verifier.py -- 감사주석 검증 도구 (PyQt6 GUI)

[동작]
  ① 시트 매칭 : 정산표/DSD와 감사조서의 순수-숫자 시트명이 일치하는 쌍 처리
                (텍스트 시트명·'숫자-...' 형태는 부속시트로 제외)
  ② 블록 감지 : 감사조서 시트에서 왼쪽 블록 / 빈열 간격 / 오른쪽 블록 자동 탐지
  ③ 복사      : 정산표 시트 내용(라벨+값)을 감사조서 왼쪽 블록에 행 위치 기준으로 덮어씀
  ④ 비교·색상 : 왼쪽 블록 숫자 vs 오른쪽 블록(감사인 작성) 숫자를 비교하여 색상 표시

[색상 기준]
  연녹  : 일치 (차이 1천원 이하)
  연황  : 소차이 (1천원 초과 ~ 1백만원 이하)
  연적  : 큰차이 (1백만원 초과)
"""

import os
import re
import sys
from datetime import datetime

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

import xlwings as xw
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFileDialog, QTextEdit, QRadioButton,
    QButtonGroup, QGroupBox, QFrame, QProgressBar,
)
from PyQt6.QtCore import QThread, pyqtSignal, Qt
from PyQt6.QtGui import QFont, QColor, QTextCursor

# ── 색상 상수 (RGB 튜플) ──────────────────────────────────────────────
COLOR_OK   = (198, 239, 206)   # 연녹  : 일치
COLOR_DIFF = (255, 235, 156)   # 연황  : 소차이
COLOR_BIG  = (255, 199, 206)   # 연적  : 큰차이
COLOR_HDR  = (189, 215, 238)   # 파랑  : 요약 헤더

THRESH_SMALL = 1        # 1천원 이하 → 일치
THRESH_BIG   = 1_000   # 1백만원(천원 단위) 초과 → 큰 차이


# ════════════════════════════════════════════════════════════════════
# SheetData — xlwings used_range 결과를 cell(r,c) 인터페이스로 감쌈
# ════════════════════════════════════════════════════════════════════

class _Cell:
    __slots__ = ('value',)
    def __init__(self, v): self.value = v


class SheetData:
    def __init__(self, data: list):
        self._d = data
        self.max_row    = len(data)
        self.max_column = max((len(r) for r in data), default=0)

    def cell(self, row: int, col: int) -> _Cell:
        try:
            v = self._d[row - 1][col - 1]
        except IndexError:
            v = None
        return _Cell(v)

    def iter_rows(self, min_row=1, max_row=None, values_only=True):
        mr = min(max_row or self.max_row, self.max_row)
        for r in range(min_row - 1, mr):
            row = self._d[r]
            yield tuple(row[c] if c < len(row) else None
                        for c in range(self.max_column))


def _read_xw_sheet(xw_sheet) -> SheetData:
    """xlwings 시트 전체를 한 번의 배치 읽기로 SheetData에 담는다."""
    vals = xw_sheet.used_range.value
    if vals is None:
        return SheetData([])
    if not isinstance(vals, list):
        return SheetData([[vals]])
    if not vals or not isinstance(vals[0], list):
        return SheetData([vals])
    return SheetData(vals)


# ════════════════════════════════════════════════════════════════════
# 헬퍼 함수
# ════════════════════════════════════════════════════════════════════

def _is_num(v) -> bool:
    return isinstance(v, (int, float)) and not isinstance(v, bool)


def _is_note_sheet(name: str) -> bool:
    """순수 숫자 시트명만 True. 텍스트·'숫자-기호' 형태는 False."""
    return bool(re.fullmatch(r'\s*\d+\s*', name))


def _note_sheets(names: list) -> list:
    """감사주석 비교 대상 시트 목록 (순수 숫자명만)."""
    return [s for s in names if _is_note_sheet(s)]


def _find_blocks_in_range(ws: SheetData, row_start: int, row_end: int):
    """
    지정 행 범위(row_start~row_end) 안에서 왼쪽 블록 끝 열(b1e)과
    오른쪽 블록 시작 열(b2s)을 탐지.
    연속 2개 이상 빈 열을 간격으로 판단.
    반환: (b1e, b2s) 또는 (None, None)
    """
    mc = ws.max_column

    def col_is_empty(c):
        return all(ws.cell(r, c).value is None for r in range(row_start, row_end + 1))

    for c in range(2, mc):
        if col_is_empty(c) and col_is_empty(c + 1):
            b1e = c - 1
            for b2s in range(c + 1, mc + 2):
                if b2s > mc:
                    return None, None
                if not col_is_empty(b2s):
                    return b1e, b2s
    return None, None


def _find_tables(ws: SheetData, col_start=1, col_end=None, min_cols=1) -> list:
    """
    빈 행으로 구분된 표 구간 목록 반환: [(row_start, row_end), ...]
    col_start~col_end 범위에 값이 있는 행을 '내용 있음'으로 판단.
    min_cols: 표로 인정하려면 최소 이 개수의 열에 값이 있어야 함.
              소스 탐지 시 2를 지정하면 열1만 있는 제목 행을 표에서 제외.
    """
    col_end = col_end or ws.max_column
    tables, in_table, t_start = [], False, None
    for r in range(1, ws.max_row + 1):
        filled = any(
            ws.cell(r, c).value is not None
            for c in range(col_start, col_end + 1)
        )
        if filled and not in_table:
            in_table, t_start = True, r
        elif not filled and in_table:
            tables.append((t_start, r - 1))
            in_table = False
    if in_table:
        tables.append((t_start, ws.max_row))
    if min_cols > 1:
        filtered = []
        for (ts, te) in tables:
            cols_with_data = set()
            for r in range(ts, te + 1):
                for c in range(col_start, col_end + 1):
                    if ws.cell(r, c).value is not None:
                        cols_with_data.add(c)
            if len(cols_with_data) >= min_cols:
                filtered.append((ts, te))
        return filtered
    return tables


def _has_border(xw_sheet, row_start: int, row_end: int,
                col_start: int, col_end: int) -> bool:
    """
    지정 범위 안에 Excel 셀 테두리가 하나라도 있으면 True.
    border_id: 7=left 8=top 9=bottom 10=right 11=inside_v 12=inside_h
    xlLineStyleNone = -4142
    """
    col_scan = min(col_end, col_start + 10)
    try:
        rng = xw_sheet.range((row_start, col_start), (row_end, col_scan))
        for bid in (9, 12, 8, 10, 11, 7):
            if rng.api.Borders(bid).LineStyle != -4142:
                return True
        return False
    except Exception:
        return True


def _trim_right_table_start(ws: SheetData, ts: int, te: int,
                            col_start: int, col_end: int) -> int:
    """
    오른쪽 블록 표 선두에서 <당기말>/<전기말> 형식의 꺾쇠 라벨 행만 건너뜁니다.
    금융자산, 구분 등 일반 헤더 행은 그대로 유지합니다.
    """
    angle_re = re.compile(r'^\s*<[^>]+>\s*$')
    col_end = min(col_end, ws.max_column)
    for r in range(ts, te + 1):
        non_none = [ws.cell(r, c).value
                    for c in range(col_start, col_end + 1)
                    if ws.cell(r, c).value is not None]
        if not non_none:
            continue
        if (len(non_none) == 1
                and isinstance(non_none[0], str)
                and angle_re.match(str(non_none[0]))):
            continue
        return r
    return ts


def _trim_table_start(ws: SheetData, ts: int, te: int,
                      min_cols: int = 2,
                      col_start: int = 1, col_end: int = None) -> int:
    """
    블록(ts~te) 내에서 처음으로 min_cols개 이상의 열에 값이 있는 행 번호를 반환.
    선두의 단일열 설명·단위 행(예: '<당기말>', '(단위:천원)')을 건너뛰어
    실제 표 헤더가 시작되는 행을 찾는다.
    """
    col_end = col_end or ws.max_column
    for r in range(ts, te + 1):
        count = sum(1 for c in range(col_start, col_end + 1)
                    if ws.cell(r, c).value is not None)
        if count >= min_cols:
            return r
    return ts


def _first_numeric_row(ws: SheetData, row_start: int, row_end: int,
                       col_start: int, col_end: int) -> int:
    """지정 범위에서 숫자 셀이 처음 나타나는 행 반환. 없으면 row_start 반환."""
    col_end = min(col_end, ws.max_column)
    for r in range(row_start, row_end + 1):
        for c in range(col_start, col_end + 1):
            if _is_num(ws.cell(r, c).value):
                return r
    return row_start


def _match_left(lefts: list, rights: list) -> list:
    """
    오른쪽 블록 표(rights)에 대응하는 왼쪽 블록 표(lefts)를 행 시작 위치 근접도로 1:1 매칭.
    매칭 실패 시 오른쪽 표 범위를 그대로 반환(fallback).
    """
    used, result = set(), []
    for rs, re in rights:
        best_i, best_d = None, float('inf')
        for i, (ls, le) in enumerate(lefts):
            if i in used:
                continue
            d = abs(ls - rs)
            if d < best_d:
                best_d, best_i = d, i
        if best_i is not None and best_d <= 10:
            used.add(best_i)
            result.append(lefts[best_i])
        else:
            result.append((rs, re))
    return result


def _filter_src_tables(xw_sheet, ws: SheetData, tables: list) -> list:
    """
    소스 표 후보에서 실제 데이터 표가 아닌 제목·설명 블록을 제거.
    전략:
      1) border 적응형: 시트 내 표에 border가 하나라도 있으면 border 없는 표를 제거.
      2) border 없는 파일(border=0개): 숫자 셀이 1개 이상 있는 표만 통과.
    이 두 단계로 순수 텍스트 제목 행을 제거하고 실제 데이터 표만 남긴다.
    """
    if not tables:
        return tables

    # border 확인
    border_flags = [_has_border(xw_sheet, t[0], t[1], 1, ws.max_column) for t in tables]
    any_border = any(border_flags)

    if any_border:
        # border가 존재하는 파일: border 없는 블록 제거
        return [t for t, b in zip(tables, border_flags) if b]
    else:
        # border 없는 파일: 숫자 셀이 1개 이상 있는 표만 통과
        result = []
        for (ts, te) in tables:
            has_num = any(
                _is_num(ws.cell(r, c).value)
                for r in range(ts, te + 1)
                for c in range(1, ws.max_column + 1)
            )
            if has_num:
                result.append((ts, te))
        return result


# ════════════════════════════════════════════════════════════════════
# 메인 검증 함수
# ════════════════════════════════════════════════════════════════════

def _count_trailing_commas(fmt: str) -> int:
    """Excel 서식문자열 첫 섹션 끝의 쉼표 수 반환 (쉼표 1개 = 표시값 ÷1,000)."""
    if not fmt:
        return 0
    first = fmt.split(';')[0].rstrip()
    count = 0
    for ch in reversed(first):
        if ch == ',':
            count += 1
        else:
            break
    return count


def _detect_unit(xw_sheet, ws: SheetData) -> float:
    """
    소스 시트 단위 자동 감지.
    1) 값 열(c≥2) 첫 금액 셀의 NumberFormat 끝 쉼표 확인:
       끝 쉼표 있음 → #,###, 계열 서식 → 실제 원단위 저장 → 0.001 반환.
    2) 폴백: 숫자 중간값 > 10,000,000 이면 원단위 → 0.001 반환.
    """
    # 1) 서식 기반 감지 (xlwings COM 호출 1회)
    try:
        found_fmt = False
        for r in range(1, min(ws.max_row + 1, 150)):
            if found_fmt:
                break
            for c in range(2, min(ws.max_column + 1, 20)):
                v = ws.cell(r, c).value
                if _is_num(v) and abs(v) > 1_000:
                    fmt = xw_sheet.range((r, c)).number_format or ''
                    if _count_trailing_commas(fmt) >= 1:
                        return 0.001   # 원단위 저장 / 천원 표시 서식
                    found_fmt = True   # 서식 확인됐으나 끝 쉼표 없음 → 폴백
                    break
    except Exception:
        pass

    # 2) 중간값 기반 폴백
    nums = []
    for r in range(1, min(ws.max_row + 1, 200)):
        for c in range(1, min(ws.max_column + 1, 30)):
            v = ws.cell(r, c).value
            if _is_num(v) and v > 0:
                nums.append(v)
    if not nums:
        return 1.0
    nums.sort()
    median = nums[len(nums) // 2]
    return 0.001 if median > 10_000_000 else 1.0


def run_verify(audit_path: str, src_path: str,
               src_label: str, unit_scale: float, log,
               progress=None, status=None) -> str:
    """
    unit_scale: 1.0(천원 그대로) or 0.001(원→천원 변환) or None(자동 감지)
    반환: 저장된 파일 경로
    """
    log(f"감사조서  : {os.path.basename(audit_path)}")
    log(f"소스({src_label}): {os.path.basename(src_path)}")
    if unit_scale != 1.0:
        log("단위 변환 : ÷1,000 (원 → 천원)")

    if status: status("Excel 엔진 시작 중...")
    app = xw.App(visible=False, add_book=False)
    app.display_alerts = False

    try:
        if status: status(f"소스 파일 로드 중...  ({os.path.basename(src_path)})")
        xw_src = app.books.open(src_path)

        # 단위 자동 감지 (unit_scale == 1.0 이어도 덮어씀)
        if status: status("소스 단위 자동 감지 중...")
        _first_num_sheet = next(
            (s for s in xw_src.sheets if _is_note_sheet(s.name)), None)
        if _first_num_sheet is not None:
            _sample = _read_xw_sheet(_first_num_sheet)
            auto_scale = _detect_unit(_first_num_sheet, _sample)
            if auto_scale != unit_scale:
                unit_scale = auto_scale
                label_unit = "원 → 천원 (자동 감지)" if unit_scale == 0.001 else "천원 그대로 (자동 감지)"
                log(f"단위 자동 감지 : {label_unit}")

        if status: status(f"감사조서 로드 중...  ({os.path.basename(audit_path)})")
        xw_audit = app.books.open(audit_path)

        src_names   = [s.name for s in xw_src.sheets]
        audit_names = [s.name for s in xw_audit.sheets]

        # 순수 숫자 시트만 대상
        note_list = _note_sheets(audit_names)
        # 소스 시트 맵: 정규화된 숫자 → 실제 시트명
        src_map = {s.strip(): s for s in src_names if _is_note_sheet(s)}

        log(f"주석 시트 : {len(note_list)}개\n{'─'*48}")

        summary      = []
        write_ops    = []   # (sname, r, c, value)   — 왼쪽 블록 복사
        color_ops    = []   # (sname, r, c, color)   — 비교 색상

        total = len(note_list)
        for idx, sname in enumerate(note_list, 1):
            if progress: progress(sname, idx, total)
            if status:   status(f"분석 중: 주석 {sname}  ({idx}/{total})")

            src_sname = src_map.get(sname.strip())
            if not src_sname:
                log(f"  [{sname:>4}] 소스 시트 없음 — 건너뜀")
                continue

            ws_src   = _read_xw_sheet(xw_src.sheets[src_sname])
            ws_audit = _read_xw_sheet(xw_audit.sheets[sname])

            # 1단계: 시트 상단 15행으로 b2s(천원단위 표 시작 열) 먼저 파악
            prelim_b1e, prelim_b2s = _find_blocks_in_range(
                ws_audit, 1, min(15, ws_audit.max_row))
            if prelim_b2s is None:
                # 폴백: 전체 행 스캔
                prelim_b1e, prelim_b2s = _find_blocks_in_range(
                    ws_audit, 1, ws_audit.max_row)
            if prelim_b2s is None:
                log(f"  [{sname:>4}] 블록 구분 열 미발견 — 건너뜀")
                continue

            # 2단계: 오른쪽 블록 기준 감사조서 표 위치 탐지 (비교 위치)
            scan_end = min(prelim_b2s + 6, ws_audit.max_column)
            audit_tables = _find_tables(ws_audit, col_start=prelim_b2s, col_end=scan_end)
            audit_tables = [(_trim_right_table_start(ws_audit, ts, te,
                                                     col_start=prelim_b2s,
                                                     col_end=scan_end), te)
                            for ts, te in audit_tables]
            # 숫자 없는 1~2행 스퍼리어스 표 제거 (총괄표 링크 등)
            audit_tables = [
                (ts, te) for ts, te in audit_tables
                if te - ts + 1 > 2 or any(
                    _is_num(ws_audit.cell(r, c).value)
                    for r in range(ts, te + 1)
                    for c in range(prelim_b2s, scan_end + 1)
                )
            ]

            # 3단계: 왼쪽 블록 표 탐지 (쓰기 위치)
            # 오른쪽 블록은 <당기말>/<전기말> 라벨 행이 앞에 있어 1행 늦게 시작하므로
            # 왼쪽 블록을 별도로 탐지해 쓰기 위치(write_r)와 비교 위치(right_r)를 분리한다.
            left_col_end = min(prelim_b2s - 3, ws_audit.max_column)
            left_raw     = _find_tables(ws_audit, col_start=1,
                                        col_end=left_col_end, min_cols=2)
            left_trimmed = [(_trim_table_start(ws_audit, ts, te, min_cols=2,
                                               col_start=1, col_end=left_col_end), te)
                            for ts, te in left_raw]
            paired_left  = _match_left(left_trimmed, audit_tables)

            src_tables = _find_tables(ws_src, min_cols=2)
            src_tables = _filter_src_tables(
                xw_src.sheets[src_sname], ws_src, src_tables)
            src_tables = [(_trim_table_start(ws_src, ts, te, min_cols=3), te)
                          for ts, te in src_tables]

            log(f"  [{sname:>4}] 소스표:{src_tables} | 감사표:{audit_tables} | 좌측표:{left_trimmed}")
            if len(src_tables) != len(audit_tables):
                log(f"  [{sname:>4}] 경고: 소스 표 {len(src_tables)}개 ≠ 감사조서 표 {len(audit_tables)}개 — 순서대로 매칭")

            filled = ok = diff_s = diff_b = 0

            for (src_s, src_e), (aud_s, aud_e), (left_s, left_e) in zip(
                    src_tables, audit_tables, paired_left):
                # 표별 b2s 정밀 탐지
                _, b2s = _find_blocks_in_range(ws_audit, aud_s, aud_e)
                if b2s is None:
                    b2s = prelim_b2s

                copy_cols = min(ws_src.max_column, b2s - 3)

                # 첫 숫자 행 기준 정렬: 헤더 행 수 차이로 인한 오프셋 제거
                src_num_s = _first_numeric_row(
                    ws_src, src_s, src_e, 1, ws_src.max_column)
                aud_num_s = _first_numeric_row(
                    ws_audit, aud_s, aud_e, b2s,
                    min(b2s + 10, ws_audit.max_column))
                data_height = min(src_e - src_num_s + 1, aud_e - aud_num_s + 1)

                log(f"  [{sname:>4}]   → src_num={src_num_s} aud_num={aud_num_s} h={data_height}")

                for i in range(data_height):
                    src_r   = src_num_s + i
                    right_r = aud_num_s + i   # 오른쪽 블록 비교 위치
                    write_r = right_r          # 왼쪽 블록: 감사조서 동일 행에 쓰기

                    for c in range(1, copy_cols + 1):
                        src_v = ws_src.cell(src_r, c).value
                        if isinstance(src_v, bool):
                            src_v = None

                        write_ops.append((sname, write_r, c, src_v))

                        if not _is_num(src_v):
                            continue

                        fill_v = round(src_v * unit_scale)
                        rc = b2s + (c - 1)
                        if rc > ws_audit.max_column:
                            continue

                        audit_raw = ws_audit.cell(right_r, rc).value
                        if not _is_num(audit_raw):
                            continue
                        audit_v = audit_raw

                        diff = abs(fill_v - audit_v)
                        if diff <= THRESH_SMALL:
                            color = COLOR_OK;   ok     += 1
                        elif diff <= THRESH_BIG:
                            color = COLOR_DIFF; diff_s += 1
                        else:
                            color = COLOR_BIG;  diff_b += 1
                            lbl = ws_src.cell(src_r, 1).value or f"R{src_r}"
                            summary.append((sname, str(lbl).strip(),
                                            fill_v, audit_v, fill_v - audit_v))

                        color_ops.append((sname, write_r, c, color))
                        filled += 1

            log(f"  [{sname:>4}] 채움 {filled:3}건  ✓{ok}  소차이 {diff_s}  큰차이 {diff_b}")

        xw_src.close()

        # ── 왼쪽 블록 값 일괄 기록 ──────────────────────────────────
        if status: status("왼쪽 블록 복사 중...")
        sheets_cache = {}

        def _get_sheet(name):
            if name not in sheets_cache:
                sheets_cache[name] = xw_audit.sheets[name]
            return sheets_cache[name]

        # 시트별로 묶어서 범위 단위 배치 쓰기
        from itertools import groupby
        for sname, ops in groupby(write_ops, key=lambda x: x[0]):
            ops = list(ops)
            if not ops:
                continue
            rows = sorted({o[1] for o in ops})
            cols = sorted({o[2] for o in ops})
            r_min, r_max = rows[0], rows[-1]
            c_min, c_max = cols[0], cols[-1]
            grid = [[None] * (c_max - c_min + 1) for _ in range(r_max - r_min + 1)]
            for _, r, c, v in ops:
                grid[r - r_min][c - c_min] = v
            _get_sheet(sname).range(
                (r_min, c_min), (r_max, c_max)
            ).value = grid

        # ── 색상 적용 ────────────────────────────────────────────────
        if status: status("색상 적용 중...")
        for sname, r, c, color in color_ops:
            _get_sheet(sname).cells(r, c).color = color

        # ── 검증요약 시트 ────────────────────────────────────────────
        log(f"\n{'─'*48}")
        log(f"큰 차이 항목 합계: {len(summary)}건")
        if status: status("검증요약 시트 생성 중...")

        SUM = '검증요약'
        if SUM in [s.name for s in xw_audit.sheets]:
            xw_audit.sheets[SUM].delete()
        ws_s = xw_audit.sheets.add(SUM, before=xw_audit.sheets[0])
        hdrs = ['주석번호', '항목', f'{src_label}값(천원)', '감사조서값(천원)', '차이(천원)']
        ws_s.range('A1').value = [hdrs]
        ws_s.range('A1:E1').api.Font.Bold = True
        ws_s.range('A1:E1').color = COLOR_HDR
        if summary:
            ws_s.range('A2').value = [list(r) for r in summary]

        # ── 저장 ────────────────────────────────────────────────────
        stem, ext = os.path.splitext(audit_path)
        out = f"{stem}_검증_{datetime.now().strftime('%m%d_%H%M')}{ext}"
        if status: status(f"저장 중...  ({os.path.basename(out)})")
        xw_audit.save(out)
        xw_audit.close()
        log(f"저장: {os.path.basename(out)}")
        return out

    finally:
        try:
            app.quit()
        except Exception:
            pass


# ════════════════════════════════════════════════════════════════════
# 백그라운드 스레드
# ════════════════════════════════════════════════════════════════════

class Worker(QThread):
    log_sig      = pyqtSignal(str)
    done_sig     = pyqtSignal(str)
    err_sig      = pyqtSignal(str)
    progress_sig = pyqtSignal(str, int, int)
    status_sig   = pyqtSignal(str)

    def __init__(self, audit, src, label, scale):
        super().__init__()
        self.audit = audit; self.src = src
        self.label = label; self.scale = scale

    def run(self):
        try:
            out = run_verify(self.audit, self.src, self.label, self.scale,
                             lambda m: self.log_sig.emit(m),
                             lambda s, i, t: self.progress_sig.emit(s, i, t),
                             lambda s: self.status_sig.emit(s))
            self.done_sig.emit(out)
        except Exception:
            import traceback
            self.err_sig.emit(traceback.format_exc())


# ════════════════════════════════════════════════════════════════════
# GUI
# ════════════════════════════════════════════════════════════════════

class NoteVerifier(QMainWindow):
    def __init__(self, default_dir: str = None):
        super().__init__()
        self.setWindowTitle("감사주석 검증 도구")
        self.resize(740, 580)
        self._dir    = default_dir or os.path.expanduser("~")
        self._worker = None
        self._build_ui()

    def _build_ui(self):
        cw = QWidget(); self.setCentralWidget(cw)
        vb = QVBoxLayout(cw)
        vb.setContentsMargins(14, 14, 14, 14); vb.setSpacing(10)

        ttl = QLabel("감사주석 검증 도구")
        ttl.setFont(QFont("맑은 고딕", 13, QFont.Weight.Bold))
        vb.addWidget(ttl)

        sep = QFrame(); sep.setFrameShape(QFrame.Shape.HLine)
        sep.setFrameShadow(QFrame.Shadow.Sunken); vb.addWidget(sep)

        grp = QGroupBox("파일 설정")
        gv  = QVBoxLayout(grp); gv.setSpacing(8)

        self._audit_lbl = self._file_row(gv, "감사조서 파일",
                                          lambda: self._pick("감사조서 선택", self._audit_lbl))
        self._src_lbl   = self._file_row(gv, "소스 파일 (정산표/DSD)",
                                          lambda: self._pick("소스 파일 선택", self._src_lbl))

        h1 = QHBoxLayout()
        h1.addWidget(QLabel("소스 유형:"))
        self._r_tb  = QRadioButton("정산표"); self._r_tb.setChecked(True)
        self._r_dsd = QRadioButton("DSD")
        bg1 = QButtonGroup(self); bg1.addButton(self._r_tb); bg1.addButton(self._r_dsd)
        h1.addWidget(self._r_tb); h1.addWidget(self._r_dsd); h1.addStretch()
        gv.addLayout(h1)

        h2 = QHBoxLayout()
        h2.addWidget(QLabel("소스 단위:"))
        self._r_1000 = QRadioButton("천원  (그대로 사용)"); self._r_1000.setChecked(True)
        self._r_won  = QRadioButton("원  (÷1,000 자동 변환)")
        bg2 = QButtonGroup(self); bg2.addButton(self._r_1000); bg2.addButton(self._r_won)
        h2.addWidget(self._r_1000); h2.addWidget(self._r_won); h2.addStretch()
        gv.addLayout(h2)

        vb.addWidget(grp)

        self._btn_run = QPushButton("▶  검증 시작")
        self._btn_run.setFont(QFont("맑은 고딕", 11, QFont.Weight.Bold))
        self._btn_run.setMinimumHeight(40)
        self._btn_run.setStyleSheet(
            "QPushButton{background:#2563EB;color:white;border-radius:6px;}"
            "QPushButton:hover{background:#1D4ED8;}"
            "QPushButton:disabled{background:#9CA3AF;}")
        self._btn_run.clicked.connect(self._run)
        vb.addWidget(self._btn_run)

        self._status_lbl = QLabel("")
        self._status_lbl.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._status_lbl.setFont(QFont("맑은 고딕", 9))
        self._status_lbl.setStyleSheet("color:#6B7280;")
        self._status_lbl.hide()
        vb.addWidget(self._status_lbl)

        self._progress = QProgressBar()
        self._progress.setTextVisible(True)
        self._progress.setMinimumHeight(18)
        self._progress.setStyleSheet(
            "QProgressBar{border:1px solid #D1D5DB;border-radius:4px;text-align:center;}"
            "QProgressBar::chunk{background:#2563EB;border-radius:3px;}")
        self._progress.hide()
        vb.addWidget(self._progress)

        grp2 = QGroupBox("실행 로그")
        gv2  = QVBoxLayout(grp2)
        self._log = QTextEdit()
        self._log.setReadOnly(True)
        self._log.setFont(QFont("Consolas", 9))
        self._log.setStyleSheet("background:#1E1E1E;color:#D4D4D4;")
        gv2.addWidget(self._log)
        vb.addWidget(grp2, stretch=1)

    def _file_row(self, parent, label_txt: str, slot) -> QLabel:
        h    = QHBoxLayout()
        lbl  = QLabel(label_txt + ":"); lbl.setFixedWidth(200)
        path = QLabel("(선택 안 됨)")
        path.setStyleSheet(
            "color:#9CA3AF;border:1px solid #E5E7EB;padding:3px 6px;border-radius:4px;")
        btn = QPushButton("파일 선택"); btn.setFixedWidth(90)
        btn.clicked.connect(slot)
        h.addWidget(lbl); h.addWidget(path, 1); h.addWidget(btn)
        parent.addLayout(h)
        return path

    def _pick(self, title: str, lbl: QLabel):
        path, _ = QFileDialog.getOpenFileName(
            self, title, self._dir, "Excel 파일 (*.xlsx *.xls)")
        if path:
            lbl.setText(path)
            lbl.setStyleSheet(
                "color:#111827;border:1px solid #D1D5DB;padding:3px 6px;border-radius:4px;")
            self._dir = os.path.dirname(path)

    def _log_ln(self, text: str, color="#D4D4D4"):
        self._log.moveCursor(QTextCursor.MoveOperation.End)
        self._log.setTextColor(QColor(color))
        self._log.insertPlainText(text + "\n")
        self._log.moveCursor(QTextCursor.MoveOperation.End)

    def _run(self):
        audit = self._audit_lbl.text()
        src   = self._src_lbl.text()
        if not os.path.isfile(audit):
            self._log_ln("감사조서 파일을 선택하세요.", "#FBBF24"); return
        if not os.path.isfile(src):
            self._log_ln("소스 파일을 선택하세요.", "#FBBF24"); return

        label = "정산표" if self._r_tb.isChecked() else "DSD"
        scale = 1.0 if self._r_1000.isChecked() else 0.001

        self._log.clear()
        self._log_ln(f"▶ {datetime.now().strftime('%H:%M:%S')} 검증 시작", "#60A5FA")
        self._btn_run.setEnabled(False)
        self._progress.setValue(0)
        self._progress.setRange(0, 0)
        self._progress.show()
        self._status_lbl.setText("Excel 엔진 시작 중...")
        self._status_lbl.setStyleSheet("color:#6B7280;")
        self._status_lbl.show()

        self._worker = Worker(audit, src, label, scale)
        self._worker.log_sig.connect(lambda m: self._log_ln(m))
        self._worker.status_sig.connect(self._status_lbl.setText)
        self._worker.progress_sig.connect(self._on_progress)
        self._worker.done_sig.connect(self._done)
        self._worker.err_sig.connect(self._err)
        self._worker.start()

    def _on_progress(self, sheet: str, cur: int, total: int):
        self._progress.setRange(0, total)
        self._progress.setValue(cur)
        self._status_lbl.setText(f"분석 중: 주석 {sheet}  ({cur} / {total})")

    def _done(self, out: str):
        self._log_ln(f"\n완료: {out}", "#4ADE80")
        self._progress.setRange(0, 1)
        self._progress.setValue(1)
        self._status_lbl.setText("✔  검증 완료")
        self._status_lbl.setStyleSheet("color:#16A34A; font-weight:bold;")
        self._btn_run.setEnabled(True)

    def _err(self, msg: str):
        self._log_ln(f"\n오류:\n{msg}", "#F87171")
        self._progress.hide()
        self._status_lbl.setText("✘  오류 발생")
        self._status_lbl.setStyleSheet("color:#DC2626; font-weight:bold;")
        self._btn_run.setEnabled(True)


# ════════════════════════════════════════════════════════════════════
# 진입점
# ════════════════════════════════════════════════════════════════════

def main():
    import argparse
    parser = argparse.ArgumentParser(add_help=False)
    parser.add_argument('--company', default=None)
    parser.add_argument('--base',    default=None)
    args, _ = parser.parse_known_args()

    default_dir = None
    if args.company:
        root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        base = (os.path.join(root, args.base, args.company)
                if args.base else os.path.join(root, args.company))
        if os.path.isdir(base):
            default_dir = base

    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = NoteVerifier(default_dir=default_dir)
    win.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
