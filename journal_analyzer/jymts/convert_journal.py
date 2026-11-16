#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
jymts 전용 분개장 변환 (더존 인쇄양식 분개장 → 엔진용 1줄 1분개 표)

원본 양식 (전기_jymts_분개장_25년.xlsx, 2026-11-16 확인):
  - 약 38행마다 페이지 머리말(제목 '분 개 장' / 기간 / 회사명 / 2단 머리글)이 반복
  - 전표 = 여러 줄. 첫 줄에만 B열 '월/일', D열 '번호'. 줄마다 차변(E 계정, F 금액) 또는
    대변(G 계정, H 금액) 한쪽만 있음. 마지막 줄(금액 없음)이 적요(E)·거래처(G) — 전표 공통
  - 날짜에 연도 없음(기간 행에서 가져옴), 전표번호는 날마다 00001부터 → 전표그룹키=일자_번호
  - 계정명 글자 사이 공백('복   리   후  생  비(판)') → 공백 제거
  - 맨 끝 '합 계' 행(차변 E, 대변 G) → 변환 결과 합계와 대사

계정별원장 (당기_jymts_계정별원장_26년.xlsx, 2026-11-16 확인):
  - 시트 1개 = 계정 1개, 시트명 '(0511)복리후생비', 9열(빈A/일자/적요/코드/빈E/거래처/차변/대변/잔액)
  - 엔진의 더존 원장 판독(_extract_duzon_style_records)은 8열·7열 양식만 알아서 거래처 열을
    차변으로 읽다 실패함(27번 오류) → 엔진이 아는 7열 양식(날짜/적요란/코드/거래처/차변/대변/잔액)
    으로 옮기고, 시트명을 '{순번}_{계정명}({코드})'로 바꾼다.
  - 계정명은 분개장과 맞춘다: 공백 제거, 코드 5xx → '(제)', 8xx → '(판)' 부여
    (분개장은 '복리후생비(제)'·'복리후생비(판)', 원장은 둘 다 '복리후생비')

엔진(main_analyzer.load_data)은 data/current·previous의 xlsx를 그대로 읽으므로, 원본은
data/raw/{current|previous}/로 옮기고 변환본만 '{원본명}_변환.xlsx'로 남긴다.
원본이 바뀌지 않았으면 다시 만들지 않는다.

사용법: main_analyzer.py jymts 실행 시 분개장 로드 전에 자동 실행됨(_run_company_converter).
        단독 실행도 가능: python journal_analyzer/jymts/convert_journal.py
"""
import datetime
import os
import re
import shutil
import sys

import openpyxl
import pandas as pd

_DIR = os.path.dirname(os.path.abspath(__file__))
_PERIODS = ('current', 'previous')
_COLS = ['전표일자', '전표번호', '전표그룹키', '계정명', '차변', '대변', '적요', '거래처명', '원본행']


def _blank(v) -> bool:
    return v is None or (isinstance(v, str) and v.strip() == '')


def _nospace(v) -> str:
    return re.sub(r'\s+', '', str(v))


def _is_raw_journal(path: str) -> bool:
    """앞쪽 10행 안에 '분개장' 제목 행이 있으면 인쇄양식 원본으로 본다."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        for i, r in enumerate(wb.worksheets[0].iter_rows(values_only=True)):
            if i >= 10:
                break
            if any(not _blank(v) and _nospace(v) == '분개장' for v in r):
                return True
    finally:
        wb.close()
    return False


def parse_journal(path: str) -> pd.DataFrame:
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    rows = [(tuple(r) + (None,) * 9)[:9] for r in wb.worksheets[0].iter_rows(values_only=True)]
    wb.close()

    year, grand, cur, vouchers = None, None, None, []
    for i, r in enumerate(rows, 1):
        b, c, d, e, f, g, h = r[1], r[2], r[3], r[4], r[5], r[6], r[7]
        bs = '' if _blank(b) else _nospace(b)
        if bs == '분개장' or bs in ('구분', '월/일'):
            continue
        if re.fullmatch(r'\d{4}\.\d{2}\.\d{2}~\d{4}\.\d{2}\.\d{2}', bs):
            year = int(bs[:4])
            continue
        if not _blank(c) and '회사명' in str(c):
            continue
        if bs == '합계':
            grand = (e, g)
            continue
        if all(_blank(v) for v in r):
            continue
        if re.fullmatch(r'\d{2}/\d{2}', bs):
            if year is None:
                raise ValueError(f'{i}행: 기간 행보다 전표가 먼저 나옴 — 연도를 알 수 없음')
            date = f'{year}-{bs[:2]}-{bs[3:]}'
            no = str(d).strip().split('.')[0].zfill(5)
            cur = {'date': date, 'no': no, 'lines': [], 'memo': (None, None)}
            vouchers.append(cur)
        elif bs:
            raise ValueError(f'{i}행: 알 수 없는 B열 값 {b!r}')
        if cur is None:
            raise ValueError(f'{i}행: 전표 시작 전에 분개 줄이 나옴')

        if not _blank(f) or not _blank(h):
            if not _blank(f):
                cur['lines'].append((i, _nospace(e), float(f), 0.0))
            if not _blank(h):
                cur['lines'].append((i, _nospace(g), 0.0, float(h)))
        else:  # 금액 없는 줄 = 적요/거래처
            cur['memo'] = (None if _blank(e) else str(e).strip(),
                           None if _blank(g) else str(g).strip())

    recs = []
    for v in vouchers:
        memo, cli = v['memo']
        for (i, acct, dr, cr) in v['lines']:
            recs.append([v['date'], v['no'], f"{v['date']}_{v['no']}", acct, dr, cr, memo, cli, i])
    df = pd.DataFrame(recs, columns=_COLS)

    # ── 검증: 파일 합계행 = 변환 합계, 전표별 차대 일치 ──
    dr_sum, cr_sum = df['차변'].sum(), df['대변'].sum()
    print(f'  전표 {len(vouchers):,}건 → 분개 {len(df):,}줄, 차변 {dr_sum:,.0f} / 대변 {cr_sum:,.0f}')
    if grand is None:
        print('  [경고] 원본에 합계 행이 없어 합계 대사를 못 함')
    else:
        for name, tot, s in (('차변', grand[0], dr_sum), ('대변', grand[1], cr_sum)):
            ok = abs(float(tot) - s) < 0.5
            print(f'  합계행 {name} {float(tot):,.0f} vs 변환 {s:,.0f} → {"일치" if ok else "불일치"}')
            if not ok:
                raise ValueError(f'{name} 합계 불일치 — 변환 중단')
    g = df.groupby('전표그룹키')[['차변', '대변']].sum()
    bad = g[(g['차변'] - g['대변']).abs() > 0.5]
    print(f'  차대 불일치 전표: {len(bad)}건' + (f' {list(bad.index[:10])}' if len(bad) else ''))
    return df


_LEDGER_RAW_HEADER = [None, '일자', '적요', '코드', None, '거래처', '차변', '대변', '잔액']
_LEDGER_OUT_HEADER = ['날짜', '적요란', '코드', '거래처', '차변', '대변', '잔액']
_LEDGER_SHEET_RE = re.compile(r'^\((\d+)\)(.+)$')


def _is_raw_ledger(path: str) -> bool:
    """첫 시트 1행이 9열 원장 머리글이면 원본 계정별원장."""
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    try:
        hdr = next(wb.worksheets[0].iter_rows(min_row=1, max_row=1, values_only=True), ())
    finally:
        wb.close()
    hdr = [None if _blank(v) else str(v).strip() for v in (tuple(hdr) + (None,) * 9)[:9]]
    return hdr == _LEDGER_RAW_HEADER


def _ledger_account_name(name: str, code: str) -> str:
    name = _nospace(name)
    c = int(code)
    if 500 <= c < 600:
        return f'{name}(제)'
    if 800 <= c < 900:
        return f'{name}(판)'
    return name


def convert_ledger(src: str, out: str) -> None:
    wb = openpyxl.load_workbook(src, read_only=True, data_only=True)
    new = openpyxl.Workbook()
    new.remove(new.active)
    n_sheet, n_row = 0, 0
    for idx, ws in enumerate(wb.worksheets):
        m = _LEDGER_SHEET_RE.match(ws.title.strip())
        if not m:
            print(f'  [경고] 시트명 형식 다름 — 건너뜀: {ws.title}')
            continue
        code, name = m.group(1), _ledger_account_name(m.group(2), m.group(1))
        suffix = f'({code})'
        safe = re.sub(r'[\\/*?:\[\]]', ' ', name)[:31 - len(str(idx)) - 1 - len(suffix)]
        title = f'{idx}_{safe}{suffix}'
        out_ws = new.create_sheet(title)
        out_ws.append(_LEDGER_OUT_HEADER)
        for i, r in enumerate(ws.iter_rows(min_row=2, values_only=True), 2):
            r = (tuple(r) + (None,) * 9)[:9]
            row = [r[1], r[2], r[3], r[5], r[6], r[7], r[8]]
            if all(_blank(v) for v in row):
                continue
            # 페이지마다 반복되는 머리말 4줄(제목 '계 정 별 원 장' / 기간 / 회사명 / 머리글) 제거
            b = '' if _blank(r[1]) else _nospace(r[1])
            if (b in ('계정별원장', '일자') or b.startswith('회사명')
                    or re.fullmatch(r'\d{4}\.\d{2}\.\d{2}~\d{4}\.\d{2}\.\d{2}', b)):
                continue
            for v in row[4:]:
                if not _blank(v) and not isinstance(v, (int, float)):
                    raise ValueError(f'{ws.title} {i}행: 금액 칸에 숫자 아닌 값 {v!r} — 변환 중단')
            out_ws.append(row)
            n_row += 1
        n_sheet += 1
    wb.close()
    new.save(out)
    print(f'  원장 {n_sheet}개 계정, {n_row:,}행 → 7열 양식(날짜/적요란/코드/거래처/차변/대변/잔액)')


def main() -> int:
    n = 0
    for period in _PERIODS:
        data_dir = os.path.join(_DIR, 'data', period)
        raw_dir = os.path.join(_DIR, 'data', 'raw', period)
        if not os.path.isdir(data_dir):
            continue
        # 1) data/{period}에 있는 원본(인쇄양식 분개장·9열 계정별원장) → data/raw/{period}로 이동
        for f in sorted(os.listdir(data_dir)):
            p = os.path.join(data_dir, f)
            if not f.endswith('.xlsx') or f.startswith('~$'):
                continue
            # 이름이 '_변환'이어도 내용이 원본 양식이면 원본으로 처리 (2026-11-16: 새 원장을
            # 변환본 이름으로 덮어써 저장 → 변환 안 된 채 엔진이 읽어 4번 오류)
            if ('분개장' in f and _is_raw_journal(p)) or ('계정별원장' in f and _is_raw_ledger(p)):
                raw_name = f[:-len('_변환.xlsx')] + '.xlsx' if f.endswith('_변환.xlsx') else f
                dest = os.path.join(raw_dir, raw_name)
                if os.path.exists(dest):   # 기존 원본은 덮어쓰지 않고 data/raw/old/로 보관
                    old_dir = os.path.join(_DIR, 'data', 'raw', 'old')
                    os.makedirs(old_dir, exist_ok=True)
                    stamp = datetime.datetime.fromtimestamp(os.path.getmtime(dest)).strftime('%Y%m%d_%H%M')
                    shutil.move(dest, os.path.join(old_dir, f'{os.path.splitext(raw_name)[0]}_{stamp}.xlsx'))
                    print(f'[{period}] 기존 원본 보관: {raw_name} → data/raw/old/')
                os.makedirs(raw_dir, exist_ok=True)
                shutil.move(p, dest)
                print(f'[{period}] 원본 이동: {f} → data/raw/{period}/{raw_name}')
        # 2) data/raw/{period} 원본 → data/{period}/{원본명}_변환.xlsx
        if not os.path.isdir(raw_dir):
            continue
        for f in sorted(os.listdir(raw_dir)):
            if not f.endswith('.xlsx') or f.startswith('~$'):
                continue
            src = os.path.join(raw_dir, f)
            out = os.path.join(data_dir, os.path.splitext(f)[0] + '_변환.xlsx')
            if os.path.isfile(out) and os.path.getmtime(out) >= os.path.getmtime(src):
                print(f'[{period}] 최신 변환본 있음 — 건너뜀: {os.path.basename(out)}')
                continue
            print(f'[{period}] 변환: {f}')
            if '계정별원장' in f:
                convert_ledger(src, out)
            else:
                parse_journal(src).to_excel(out, index=False)
            print(f'  저장: data/{period}/{os.path.basename(out)}')
            n += 1
    print(f'완료 — 새로 변환 {n}개')
    return 0


if __name__ == '__main__':
    sys.exit(main())
