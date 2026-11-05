"""
kindus 전용 전처리 (더존 위하고)

1) 계정별원장 .xls → .xlsx 변환
   위하고에서 내려받은 계정별원장은 구형 .xls(계정 1개 = 시트 1개, 시트명 '22_감가상각누계액(20300)')
   라 main_analyzer(openpyxl, .xlsx 전용)가 읽지 못한다. 같은 폴더에 '{원본명}_변환.xlsx'로
   옮겨 쓴다. 시트 안 제목 3줄은 버리고 머리글을 공용 더존 7열 양식
   (날짜/적요란/코드/거래처/차변/대변/잔액)으로 맞춘다. 원본이 바뀌지 않았으면 다시 만들지 않는다.

2) 같은 이름·다른 코드 계정 구분 (2026-11-05 blue sky 요청)
   위하고는 감가상각누계액·대손충당금을 자산별로 나누지 않고 같은 이름에 코드만 다르게 쓴다
   (감가상각누계액 203/205/207/209/211/213/220, 대손충당금 109/111). 엔진은 계정명으로만
   집계하므로 그대로 두면 27번 롤포워드·4번 TB_TO_GL 등에서 전부 합쳐진다.
   → 원장에서 바로 앞 시트(=차감 대상 자산, 예: 건물(20200) 다음 감가상각누계액(20300))
     이름을 붙여 '감가상각누계액(건물)', '대손충당금(외상매출금)'으로 바꾸고, 분개장도
     계정과목코드로 같은 이름을 쓴다.
   원장 비용계정은 '복리후생비(51100)'·'복리후생비(81100)'처럼 제조/판관 구분이 없어 분개장
   ('복리후생비(제)'·'복리후생비(판)')과 이름이 다르다 → 원장 시트명을 분개장의 같은 코드
   계정명으로 맞춘다(분개장에 없는 코드는 이름이 겹칠 때만 5xx=(제), 8xx=(판) 부여).
"""
import os
import re

import openpyxl
import pandas as pd
import xlrd

_DIR = os.path.dirname(os.path.abspath(__file__))
_LEDGER_DIRS = [os.path.join(_DIR, 'data', 'current'), os.path.join(_DIR, 'data', 'previous')]
_CONTRA_NAMES = ('감가상각누계액', '대손충당금')
_SHEET_RE = re.compile(r'^(\d+)_(.+)\((\d+)\)\s*$')
_HEADER = ['날짜', '적요란', '코드', '거래처', '차변', '대변', '잔액']

CODE_COL, ACCT_COL = '계정과목코드', '계정과목'


def _ledger_sources() -> list:
    out = []
    for d in _LEDGER_DIRS:
        if os.path.isdir(d):
            out += [os.path.join(d, f) for f in sorted(os.listdir(d))
                    if f.endswith('.xls') and '계정별원장' in f and not f.startswith('~$')]
    return out


def _parse_sheets(sheet_names: list) -> list:
    """[(시트명, 순번, 원장계정명, 5자리코드)]"""
    out = []
    for s in sheet_names:
        m = _SHEET_RE.match(s)
        if m:
            out.append((s, m.group(1), m.group(2).strip(), m.group(3)))
    return out


def _final_names(parsed: list, journal_names: dict) -> dict:
    """5자리 코드 → 엔진에서 쓸 계정명."""
    base_count = {}
    for _, _, name, _ in parsed:
        base_count[name] = base_count.get(name, 0) + 1
    out, prev_name = {}, None
    for _, _, name, code in parsed:
        c3 = int(code) // 100
        if name in _CONTRA_NAMES and prev_name:
            final = f'{name}({prev_name})'
        elif c3 in journal_names:
            final = journal_names[c3]
        elif base_count[name] > 1 and 500 <= c3 < 600:
            final = f'{name}(제)'
        elif base_count[name] > 1 and 800 <= c3 < 900:
            final = f'{name}(판)'
        else:
            final = name
        out[code] = final
        if name not in _CONTRA_NAMES:
            prev_name = final
    return out


def _sheet_title(idx: str, name: str, code: str) -> str:
    suffix = f'({code})'
    room = 31 - len(idx) - 1 - len(suffix)
    name = re.sub(r'[\\/*?:\[\]]', ' ', name)[:room]
    return f'{idx}_{name}{suffix}'


def _convert(src: str, names: dict, parsed: list) -> str:
    out_path = os.path.splitext(src)[0] + '_변환.xlsx'
    if os.path.isfile(out_path) and os.path.getmtime(out_path) >= os.path.getmtime(src):
        return out_path
    book = xlrd.open_workbook(src)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    for sname, idx, _, code in parsed:
        sh = book.sheet_by_name(sname)
        ws = wb.create_sheet(_sheet_title(idx, names[code], code))
        ws.append(_HEADER)
        start = next((r + 1 for r in range(sh.nrows)
                      if str(sh.cell_value(r, 0)).strip() == '날짜'), None)
        if start is None:
            continue
        for r in range(start, sh.nrows):
            row = [sh.cell_value(r, c) if c < sh.ncols else None for c in range(7)]
            row = [None if v == '' else v for v in row]
            ws.append(row)
    wb.save(out_path)
    print(f'  [kindus/preprocess] 계정별원장 변환: {os.path.basename(src)} → '
          f'{os.path.basename(out_path)} ({len(parsed)}개 계정)')
    return out_path


def preprocess(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    if CODE_COL not in df.columns or ACCT_COL not in df.columns:
        print('  [kindus/preprocess] 계정과목코드/계정과목 컬럼 없음 → 전처리 건너뜀')
        return df

    # 위하고 분개장 맨 끝 '합계 (건수 : 8,503 )' 행 제거 — 그대로 두면 금액이 문자열
    # '157,276,895,056'이라 공통 전처리가 숫자로 바꿔 차대변 합계가 2배가 된다(전기 25년 파일).
    # 지우기 전에 합계행 금액 = 전표 합계인지 검증해서 결과를 출력한다.
    total_mask = df['전표일자'].astype(str).str.strip().str.startswith('합계')
    if total_mask.any():
        def _num(s):
            return pd.to_numeric(s.astype(str).str.replace(',', '', regex=False), errors='coerce').fillna(0)
        for i in df.index[total_mask]:
            label = df.at[i, '구분'] if '구분' in df.columns else ''
            body = df[~total_mask & ((df['구분'] == label) if '구분' in df.columns else True)]
            for col in ('차변', '대변'):
                tot, body_sum = float(_num(df.loc[[i], col]).iat[0]), float(_num(body[col]).sum())
                ok = '일치' if abs(tot - body_sum) < 1 else f'불일치(차이 {tot - body_sum:,.0f})'
                print(f'  [kindus/preprocess] {label} 분개장 합계행 {col} {tot:,.0f} vs 전표합계 {body_sum:,.0f} → {ok}')
        df = df[~total_mask].copy()

    codes = pd.to_numeric(df[CODE_COL], errors='coerce')
    acct = df[ACCT_COL].astype(str).str.strip()
    # 분개장의 코드 → 계정명 (당기 우선; 대체 대상 계정은 제외)
    journal_names = {}
    for label in ('당기', '전기'):
        sub = df[(df.get('구분') == label) & codes.notna() & ~acct.isin(_CONTRA_NAMES)] \
            if '구분' in df.columns else df[codes.notna() & ~acct.isin(_CONTRA_NAMES)]
        for c, n in zip(codes[sub.index].astype(int), acct[sub.index]):
            journal_names.setdefault(c, n)

    contra_map = {}   # 3자리 코드 → 바뀐 이름 (분개장 적용용, 당기 원장 우선)
    for src in _ledger_sources():
        parsed = _parse_sheets(xlrd.open_workbook(src, on_demand=True).sheet_names())
        names = _final_names(parsed, journal_names)
        _convert(src, names, parsed)
        for _, _, name, code in parsed:
            if name in _CONTRA_NAMES:
                contra_map.setdefault(int(code) // 100, names[code])

    mask = acct.isin(_CONTRA_NAMES) & codes.notna()
    new = codes[mask].astype(int).map(contra_map)
    missing = sorted(set(codes[mask][new.isna()].astype(int)))
    df.loc[new.dropna().index, ACCT_COL] = new.dropna()
    print(f'  [kindus/preprocess] 감가상각누계액·대손충당금 계정명 구분: {int(new.notna().sum()):,}행 '
          f'({len(set(new.dropna()))}개 계정)'
          + (f' — 원장에서 짝을 못 찾은 코드 {missing}' if missing else ''))
    return df
