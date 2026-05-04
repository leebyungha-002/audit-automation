#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""매핑 리스트 기반 데이터 주입 엔진
Usage: python data_injector.py <company_name>

기준 폴더: ../<company>/감사조서/
- 매핑 파일:  ../<company>/감사조서/<company>_mapping_list*.xlsx
- 대상 조서:  ../<company>/감사조서/ (매핑의 '대상 조서 파일명' 키워드로 탐색)
- 소스 데이터: ../<company>/results/ → raw_data/ → <company>/ 순서로 탐색
"""

import sys
import os
import re
import datetime
from collections import defaultdict
from io import BytesIO
from openpyxl import load_workbook
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils import column_index_from_string

# Windows 콘솔 한글·특수문자 출력 보장
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')


# ─── 파일명 정규화 ────────────────────────────────────────────────────────────

def _normalize(fname_no_ext):
    """YYYYMMDD(8자리) 날짜 토큰만 제거 후 연속 _ 정리."""
    s = re.sub(r'_\d{8}(?=_|$)', '', fname_no_ext)
    return re.sub(r'_+', '_', s).strip('_')


def _keyword_matches(keyword, normalized_fname):
    """keyword가 정규화된 파일명 내에 _ 경계로 정확히 포함되는지 확인.

    keyword='dae_il_외상매출금'
      normalized='dae_il_외상매출금'             → True  (정확일치)
      normalized='dae_il_외상매출금_상세'         → True  (suffix 허용)
      normalized='dae_il_벤포드_외상매출금'       → False (앞에 다른 토큰)
    """
    text    = '_' + normalized_fname + '_'
    pattern = '_' + re.escape(keyword) + '_'
    return bool(re.search(pattern, text))


# ─── 파일 탐색 ────────────────────────────────────────────────────────────────

def find_file_by_keyword(directories, keyword, exclude_suffixes=None):
    """keyword를 파일명(날짜 정규화 후)에서 _ 경계로 탐색.

    여러 개 발견 시 가장 최근 수정 파일 반환.
    """
    if isinstance(directories, str):
        directories = [directories]
    if exclude_suffixes is None:
        exclude_suffixes = ['~$', '_updated']

    matches = []
    for directory in directories:
        if not os.path.isdir(directory):
            continue
        for fname in sorted(os.listdir(directory)):
            if any(ex in fname for ex in exclude_suffixes):
                continue
            if not fname.lower().endswith('.xlsx'):
                continue
            normalized = _normalize(os.path.splitext(fname)[0])
            if _keyword_matches(keyword, normalized):
                matches.append(os.path.join(directory, fname))

    if not matches:
        return None
    if len(matches) == 1:
        return matches[0]

    matches.sort(key=lambda p: os.path.getmtime(p), reverse=True)
    names = [os.path.basename(p) for p in matches]
    print(f"    [주의] '{keyword}' 키워드로 {len(matches)}개 파일 발견:")
    for n in names:
        print(f"           {n}")
    print(f"           → 최근 파일 선택: {names[0]}")
    return matches[0]


# ─── 시트 탐색 ────────────────────────────────────────────────────────────────

def resolve_sheet(sheetnames, keyword):
    """정확히 일치 → keyword 포함 첫 번째 시트 순으로 탐색."""
    if keyword in sheetnames:
        return keyword
    matched = [s for s in sheetnames if keyword in s]
    return matched[0] if matched else None


# ─── 셀 좌표 / 범위 파싱 ─────────────────────────────────────────────────────

def _parse_cell(cell_ref):
    """'A7' → (row=7, col=1)  /  대소문자 무관."""
    m = re.match(r'^([A-Za-z]+)(\d+)$', cell_ref.strip())
    if not m:
        raise ValueError(f"잘못된 셀 좌표: {cell_ref}")
    return int(m.group(2)), column_index_from_string(m.group(1).upper())


def _parse_range(range_str):
    """'B2:C13' → (min_row=2, min_col=2, max_row=13, max_col=3)."""
    parts = range_str.strip().upper().split(':')
    if len(parts) != 2:
        raise ValueError(f"잘못된 범위: {range_str}  (형식 예: B2:C13)")
    min_row, min_col = _parse_cell(parts[0])
    max_row, max_col = _parse_cell(parts[1])
    return min_row, min_col, max_row, max_col


# ─── 데이터 주입 ──────────────────────────────────────────────────────────────

def inject_data(ws_src, ws_tgt, start_cell, src_range=None):
    """소스 시트 데이터를 대상 시트의 start_cell 부터 값만 주입 (서식·수식 보존).

    src_range 지정 ('B2:C13') : 해당 영역만 추출하여 주입
    src_range 미지정 (None)   : 소스 시트 used range 전체 주입
    행·열 구조(Matrix)는 그대로 유지.  반환: 주입된 셀 수
    """
    start_row, start_col = _parse_cell(start_cell)

    if src_range:
        min_row, min_col, max_row, max_col = _parse_range(src_range)
        src_rows = ws_src.iter_rows(
            min_row=min_row, max_row=max_row,
            min_col=min_col, max_col=max_col,
            values_only=True,
        )
    else:
        src_rows = ws_src.iter_rows(values_only=True)

    count = 0
    for r_idx, row in enumerate(src_rows):
        for c_idx, value in enumerate(row):
            if value is not None:
                ws_tgt.cell(row=start_row + r_idx,
                            column=start_col + c_idx).value = value
                count += 1
    return count


# ─── 매핑 파일 로드 ──────────────────────────────────────────────────────────

def load_mapping(mapping_path):
    """<회사>_mapping_list*.xlsx 를 읽어 매핑 행 리스트 반환.

    컬럼 순서 (A~H):
      A 계정과목(label) / B 소스파일명(src_kw) / C 소스시트(src_sheet)
      D 소스 데이터 범위(src_range, 선택 — 예: B2:C13)
      E 대상파일명(tgt_kw) / F 대상시트(tgt_sheet) / G 시작셀(start_cell)
      H 비고(remarks, 선택 — 예: PIVOT_AGING)
    """
    wb = load_workbook(mapping_path, data_only=True)
    ws = wb.active
    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not any(row):
            continue
        padded = list(row) + [None] * 8
        label, src_kw, src_sheet, src_range, tgt_kw, tgt_sheet, start_cell, remarks = padded[:8]
        if not src_kw or not tgt_kw or not start_cell:
            continue
        rows.append({
            'label':      str(label      or '').strip(),
            'src_kw':     str(src_kw    ).strip(),
            'src_sheet':  str(src_sheet ).strip(),
            'tgt_kw':     str(tgt_kw    ).strip(),
            'tgt_sheet':  str(tgt_sheet ).strip(),
            'start_cell': str(start_cell).strip().upper(),
            'src_range':  str(src_range ).strip().upper() if src_range else '',
            'remarks':    str(remarks   ).strip().upper() if remarks else '',
        })
    return rows


# ─── Aging 피벗 ──────────────────────────────────────────────────────────────

def _parse_month_key(date_val):
    """날짜값(datetime / Excel 숫자 / 문자열)에서 'YYYY-MM' 문자열 반환. 실패 시 None."""
    if isinstance(date_val, (datetime.datetime, datetime.date)):
        return f"{date_val.year}-{date_val.month:02d}"
    if isinstance(date_val, (int, float)):
        try:
            dt = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=int(date_val))
            return f"{dt.year}-{dt.month:02d}"
        except Exception:
            return None
    if isinstance(date_val, str):
        s = date_val.strip()
        if len(s) >= 8 and s[:8].isdigit():          # '20250115...'
            return f"{s[:4]}-{s[4:6]}"
        for sep in ('-', '/'):
            if sep in s:
                parts = s.split(sep)
                if len(parts) >= 2:
                    return f"{parts[0].zfill(4)}-{parts[1].zfill(2)}"
    return None


def build_pivot_aging(ws_src):
    """소스 시트에서 거래처명 × 월별 차변금액 피벗 테이블을 생성한다.

    Returns (headers, data_rows):
      headers    = ['거래처명', '2025-01', ..., '합계']
      data_rows  = [['거래처A', 100000, None, ..., 100000], ..., ['합계', ...]]
    """
    # ── 헤더 행 탐색 ─────────────────────────────────────────────────────
    def _find_col(header_row, *keywords):
        for i, h in enumerate(header_row):
            cell = str(h or '')
            if any(kw in cell for kw in keywords):
                return i
        return None

    all_rows = list(ws_src.iter_rows(values_only=True))
    header_row_idx = None
    header_row = None
    for i, row in enumerate(all_rows):
        if _find_col(row, '거래처') is not None:
            header_row_idx = i
            header_row = row
            break

    if header_row is None:
        raise ValueError("소스 시트에서 '거래처명' 헤더를 찾을 수 없습니다.")

    col_cust = _find_col(header_row, '거래처')
    col_date = _find_col(header_row, '전표날짜', '날짜', '일자')
    col_amt  = _find_col(header_row, '차변금액', '차변', '금액')

    missing = [name for name, idx in [('거래처명', col_cust), ('전표날짜', col_date), ('차변금액', col_amt)] if idx is None]
    if missing:
        raise ValueError(f"필수 컬럼을 찾을 수 없습니다: {', '.join(missing)}")

    # ── 데이터 집계 ───────────────────────────────────────────────────────
    pivot  = defaultdict(lambda: defaultdict(float))
    months = set()

    for row in all_rows[header_row_idx + 1:]:
        cust = row[col_cust] if len(row) > col_cust else None
        date = row[col_date] if len(row) > col_date else None
        amt  = row[col_amt]  if len(row) > col_amt  else None

        if cust is None:
            continue

        month_key = _parse_month_key(date)
        if month_key is None:
            continue

        if isinstance(amt, str):
            amt = re.sub(r'[,원\s]', '', amt)
        try:
            amt_val = float(amt) if amt is not None else 0.0
        except (ValueError, TypeError):
            amt_val = 0.0

        pivot[str(cust).strip()][month_key] += amt_val
        months.add(month_key)

    if not pivot:
        raise ValueError("피벗 데이터 없음 — 소스 시트가 비어있거나 컬럼 매칭 실패.")

    sorted_months = sorted(months)
    headers = ['거래처명'] + sorted_months + ['합계']

    data_rows = []
    for cust in sorted(pivot.keys()):
        row_vals = [cust]
        row_total = 0.0
        for m in sorted_months:
            v = pivot[cust].get(m, 0.0)
            row_vals.append(v if v else None)
            row_total += v
        row_vals.append(row_total if row_total else None)
        data_rows.append(row_vals)

    col_totals = []
    grand_total = 0.0
    for m in sorted_months:
        ct = sum(pivot[c].get(m, 0.0) for c in pivot)
        col_totals.append(ct if ct else None)
        grand_total += ct
    data_rows.append(['합계'] + col_totals + [grand_total if grand_total else None])

    return headers, data_rows


def inject_pivot_aging(ws_src, wb_tgt, tgt_sheet_name, start_cell):
    """피벗 Aging 테이블을 대상 워크북의 tgt_sheet_name 시트에 주입한다.

    시트가 없으면 새로 생성한다. 반환값: 주입된 데이터 행 수.
    """
    headers, data_rows = build_pivot_aging(ws_src)

    if tgt_sheet_name in wb_tgt.sheetnames:
        ws_aging = wb_tgt[tgt_sheet_name]
    else:
        ws_aging = wb_tgt.create_sheet(title=tgt_sheet_name)
        print(f'    [Aging] 시트 신규 생성: {tgt_sheet_name}')

    start_row, start_col = _parse_cell(start_cell)

    for c_idx, h in enumerate(headers):
        ws_aging.cell(row=start_row, column=start_col + c_idx).value = h

    for r_idx, row in enumerate(data_rows, start=1):
        for c_idx, val in enumerate(row):
            ws_aging.cell(row=start_row + r_idx, column=start_col + c_idx).value = val

    return len(data_rows)


# ─── 이미지 복사 ─────────────────────────────────────────────────────────────

def inject_image(ws_src, ws_tgt, start_cell):
    """소스 시트의 첫 번째 이미지를 대상 시트의 start_cell 위치에 복사한다.

    소스에 이미지가 없으면 로그만 남기고 0을 반환한다.
    주의: 소스 워크북은 반드시 read_only=False 로 열어야 _images 가 채워진다.
    """
    if not getattr(ws_src, '_images', None):
        print('    [MOVE_IMAGE] 소스 시트에 이미지 없음 — 건너뜀')
        return 0

    src_img = ws_src._images[0]
    img_bytes = src_img._data()
    new_img = XLImage(BytesIO(img_bytes))
    new_img.anchor = start_cell
    ws_tgt.add_image(new_img)
    return 1


# ─── 경로 헬퍼 ───────────────────────────────────────────────────────────────

def updated_path(original_path):
    """파일명 뒤에 _updated 를 붙인 경로 반환 (이미 있으면 그대로)."""
    base, ext = os.path.splitext(original_path)
    return original_path if base.endswith('_updated') else f'{base}_updated{ext}'


# ─── 메인 ────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print('Usage: python data_injector.py <company_name>')
        sys.exit(1)

    company     = sys.argv[1]
    script_dir  = os.path.dirname(os.path.abspath(__file__))
    company_dir = os.path.normpath(os.path.join(script_dir, '..', company))
    audit_dir   = os.path.join(company_dir, '감사조서')
    results_dir = os.path.join(company_dir, 'results')
    raw_dir     = os.path.join(company_dir, 'raw_data')

    print(f'[{company}] 데이터 주입 엔진 시작')
    print(f'  감사조서 폴더 : {audit_dir}')
    print(f'  소스(results) : {results_dir}')
    print(f'  소스(raw_data): {raw_dir}')

    # ── 1. 매핑 파일 탐색 ────────────────────────────────────────────────────
    mapping_path = find_file_by_keyword(audit_dir, f'{company}_mapping_list')
    if not mapping_path:
        print(f'\n[오류] 매핑 파일 없음. 키워드: {company}_mapping_list  폴더: {audit_dir}')
        sys.exit(1)
    print(f'  매핑 파일     : {os.path.basename(mapping_path)}')

    # ── 2. 매핑 읽기 ─────────────────────────────────────────────────────────
    mapping_rows = load_mapping(mapping_path)
    print(f'  매핑 항목 수  : {len(mapping_rows)}건\n')

    # ── 3. 대상 워크북 캐시 (동일 파일 중복 로드 방지) ──────────────────────
    tgt_book_cache = {}   # real_path → Workbook
    tgt_path_cache = {}   # keyword   → real_path

    errors  = []
    success = 0

    # ── 4. 매핑 처리 ─────────────────────────────────────────────────────────
    for row in mapping_rows:
        label      = row['label']
        src_kw     = row['src_kw']
        src_sheet  = row['src_sheet']
        tgt_kw     = row['tgt_kw']
        tgt_sheet  = row['tgt_sheet']
        start_cell = row['start_cell']
        src_range  = row['src_range']
        remarks    = row['remarks']

        mode_tag = f' [{remarks}]' if remarks else ''
        print(f'  [{label}]{mode_tag} {src_kw}!{src_sheet} → {tgt_kw}!{tgt_sheet} @ {start_cell}')

        # ── 소스 파일 탐색 ─────────────────────────────────────────────────
        src_path = find_file_by_keyword([results_dir, raw_dir, company_dir], src_kw)
        if not src_path:
            msg = f'소스 파일 없음: {src_kw}'
            print(f'    [오류] {msg}')
            errors.append(f'[{label}] {msg}')
            continue
        print(f'    매칭 성공 (소스) : {src_kw}')
        print(f'                    → {os.path.relpath(src_path, company_dir)}')

        # ── 소스 시트 로드 ─────────────────────────────────────────────────
        try:
            if remarks == 'MOVE_IMAGE':
                # read_only 모드에서는 ws._images 가 채워지지 않으므로 full 모드로 열기
                wb_src = load_workbook(src_path, data_only=True)
            else:
                wb_src = load_workbook(src_path, data_only=True, read_only=True)
        except Exception as e:
            msg = f'소스 파일 오픈 실패: {e}'
            print(f'    [오류] {msg}')
            errors.append(f'[{label}] {msg}')
            continue

        resolved_src = resolve_sheet(wb_src.sheetnames, src_sheet)
        if not resolved_src:
            msg = f'소스 시트 없음: {src_sheet}  (파일: {os.path.basename(src_path)})'
            print(f'    [오류] {msg}')
            errors.append(f'[{label}] {msg}')
            wb_src.close()
            continue
        if resolved_src != src_sheet:
            print(f'    시트 매칭 (소스) : {src_sheet} → {resolved_src}')
        ws_src = wb_src[resolved_src]

        # ── 대상 파일 탐색 (캐시) ─────────────────────────────────────────
        if tgt_kw not in tgt_path_cache:
            tgt_path = find_file_by_keyword(audit_dir, tgt_kw)
            if not tgt_path:
                msg = f'대상 조서 파일 없음: {tgt_kw}'
                print(f'    [오류] {msg}')
                errors.append(f'[{label}] {msg}')
                wb_src.close()
                continue
            tgt_path_cache[tgt_kw] = tgt_path
            print(f'    매칭 성공 (대상) : {tgt_kw}')
            print(f'                    → {os.path.relpath(tgt_path, company_dir)}')
        else:
            tgt_path = tgt_path_cache[tgt_kw]

        # ── 대상 워크북 로드 (캐시) ───────────────────────────────────────
        if tgt_path not in tgt_book_cache:
            try:
                tgt_book_cache[tgt_path] = load_workbook(tgt_path)
            except Exception as e:
                msg = f'대상 파일 오픈 실패: {e}'
                print(f'    [오류] {msg}')
                errors.append(f'[{label}] {msg}')
                wb_src.close()
                continue

        wb_tgt = tgt_book_cache[tgt_path]

        # ── 대상 시트 확인 ────────────────────────────────────────────────
        resolved_tgt = resolve_sheet(wb_tgt.sheetnames, tgt_sheet)
        if not resolved_tgt:
            msg = f'대상 시트 없음: {tgt_sheet}  (파일: {os.path.basename(tgt_path)})'
            print(f'    [오류] {msg}')
            errors.append(f'[{label}] {msg}')
            wb_src.close()
            continue
        if resolved_tgt != tgt_sheet:
            print(f'    시트 매칭 (대상) : {tgt_sheet} → {resolved_tgt}')
        ws_tgt = wb_tgt[resolved_tgt]

        # ── 데이터 주입 ───────────────────────────────────────────────────
        try:
            if remarks == 'PIVOT_AGING':
                print(f'    [Aging] 피벗 생성 → {tgt_sheet} @ {start_cell}')
                injected = inject_pivot_aging(ws_src, wb_tgt, tgt_sheet, start_cell)
                success += 1
                print(f'    [완료] 피벗 {injected}행 주입')
            elif remarks == 'MOVE_IMAGE':
                print(f'    [Image] 이미지 복사 → {tgt_sheet} @ {start_cell}')
                injected = inject_image(ws_src, ws_tgt, start_cell)
                success += 1
                print(f'    [완료] 이미지 {injected}개 복사')
            else:
                if src_range:
                    print(f'    소스 범위 지정 : {src_range}')
                injected = inject_data(ws_src, ws_tgt, start_cell, src_range or None)
                success += 1
                print(f'    [완료] {injected}개 셀 주입')
        except Exception as e:
            msg = f'데이터 주입 오류: {e}'
            print(f'    [오류] {msg}')
            errors.append(f'[{label}] {msg}')

        wb_src.close()

    # ── 5. 결과 저장 ─────────────────────────────────────────────────────────
    print('\n─── 저장 ───')
    for tgt_path, wb in tgt_book_cache.items():
        out_path = updated_path(tgt_path)
        try:
            wb.save(out_path)
            print(f'  저장 완료: {os.path.relpath(out_path, company_dir)}')
        except Exception as e:
            print(f'  [오류] 저장 실패 ({os.path.basename(tgt_path)}): {e}')

    # ── 6. 요약 ──────────────────────────────────────────────────────────────
    print('\n─── 작업 요약 ───')
    print(f'  성공 : {success}/{len(mapping_rows)}건')
    if errors:
        print(f'  오류 ({len(errors)}건):')
        for err in errors:
            print(f'    - {err}')
    else:
        print('  오류 없음')


if __name__ == '__main__':
    main()
