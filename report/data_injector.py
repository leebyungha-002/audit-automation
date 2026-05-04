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
from openpyxl import load_workbook

from mapping import find_file_by_keyword, resolve_sheet, load_mapping, inject_data, updated_path

# Windows 콘솔 한글·특수문자 출력 보장
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')


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

        print(f'  [{label}] {src_kw}!{src_sheet} → {tgt_kw}!{tgt_sheet} @ {start_cell}')

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
        src_range = row.get('src_range', '')
        if src_range:
            print(f'    소스 범위 지정 : {src_range}')
        try:
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
