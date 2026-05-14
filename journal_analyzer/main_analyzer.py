#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
journal_analyzer/main_analyzer.py
분개장 분析 자동화 메인 스크립트

폴더 구조:
  journal_analyzer/
  ├── main_analyzer.py
  └── {company_name}/
      ├── task_list_{company_name}.xlsx  ← 실행 제어 + 분析 파라미터
      ├── raw_data/                       ← 분개장 파일 (.xlsx)
      └── results/                        ← 분析 결과 저장

task_list 시트 구조:
  분析목록       : 분析번호 / 분析명 / 설명 / 실행여부  ← 분析 타입별 ON/OFF
  벤포드분析     : 계정과목 / 금액열 / 실행여부 / 비고
  거액전표       : 계정과목 / 금액열 / 기준백분위(%) / 실행여부 / 비고
  상세검색       : 작업명 / 계정과목 / 거래처명 / 적요 / 금액유형 / 최소금액 / 최대금액 / 실행여부
  라운드넘버     : 계정과목 / 금액열 / 최소금액 / 실행여부 / 비고
  (공휴일전표, 중복전표는 파라미터 시트 없이 분析목록 실행여부만으로 제어)

실행:
  python main_analyzer.py graphy
  python main_analyzer.py          ← 입력창 표시
"""

import sys
import os
import glob
import argparse
import warnings

import pandas as pd

warnings.filterwarnings('ignore', category=pd.errors.DtypeWarning)

BASE_DIR          = os.path.dirname(os.path.abspath(__file__))
TASK_MASTER_SHEET = '분析목록'


# ═══════════════════════════════════════════════════════════════════════════════
# 유틸
# ═══════════════════════════════════════════════════════════════════════════════

def _nv(val, blank_vals=('(전체)', 'nan', 'none', '')) -> str:
    """None / NaN / '(전체)' 등 빈값 처리 후 strip된 문자열 반환. 빈값이면 ''."""
    s = str(val).strip()
    return '' if s.lower() in blank_vals else s


def _safe_float(val, default: float) -> float:
    """NaN·None·빈문자열이면 default, 아니면 float 변환."""
    try:
        f = float(val)
        return default if pd.isna(f) else f
    except (TypeError, ValueError):
        return default


def _filter_df(df: pd.DataFrame, params: dict):
    """
    params의 '계정과목' / '금액열' 기준으로 df 를 필터링한다.
      계정과목 : 비어있거나 '(전체)' → 전체 행
                그 외 → str.contains 부분일치
      금액열   : '차변' / '대변' → 해당 열만 반환
                그 외('양방향' 등) → ['차변', '대변'] 중 df에 존재하는 열 반환
    반환: (filtered_df, amount_col_list)
    """
    acct   = _nv(params.get('계정과목', ''))
    col_f  = _nv(params.get('금액열', ''), blank_vals=('nan', 'none', ''))

    sub = df.copy()
    if acct and '계정과목' in sub.columns:
        sub = sub[sub['계정과목'].astype(str).str.contains(acct, na=False)]

    if col_f in ('차변', '대변'):
        amount_cols = [col_f] if col_f in sub.columns else []
    else:
        amount_cols = [c for c in ('차변', '대변') if c in sub.columns]

    return sub, amount_cols


def _result_key(task_no: int, task_name: str, params: dict) -> str:
    """결과 시트명(31자 이내) 생성."""
    job  = _nv(params.get('작업명', ''))
    acct = _nv(params.get('계정과목', ''))
    col  = _nv(params.get('금액열', ''), blank_vals=('nan', 'none', ''))

    if job:
        parts = [str(task_no), job]
    else:
        parts = [str(task_no), task_name]
        if acct:
            parts.append(acct)
        if col:
            parts.append(col)

    return '_'.join(parts)[:31]


# ═══════════════════════════════════════════════════════════════════════════════
# 1. 경로 세팅
# ═══════════════════════════════════════════════════════════════════════════════

def resolve_paths(company_name: str) -> dict:
    """
    journal_analyzer/ 하위 고객사 폴더 기준 경로 세팅.
    BASE_DIR       = journal_analyzer/
    company_dir    = journal_analyzer/{company_name}/
    """
    company_dir = os.path.join(BASE_DIR, company_name)
    return {
        'company_dir': company_dir,
        'task_list':   os.path.join(company_dir, f'task_list_{company_name}.xlsx'),
        'raw_data':    os.path.join(company_dir, 'raw_data'),
        'output':      os.path.join(company_dir, 'results'),
    }


# ═══════════════════════════════════════════════════════════════════════════════
# 2. Task List 읽기
# ═══════════════════════════════════════════════════════════════════════════════

def load_active_tasks(task_list_path: str) -> list:
    """
    '분析목록' 시트에서 실행여부=Y/O 인 (번호, 명) 튜플 리스트 반환.
    '분析목록' 시트가 없으면 [분析번호 / 분析명 / 실행여부] 컬럼을 가진 시트를 자동 탐색.
    """
    if not os.path.isfile(task_list_path):
        raise FileNotFoundError(f"task_list 파일이 없습니다: {task_list_path}")

    xl       = pd.ExcelFile(task_list_path)
    df, hit  = None, None

    if TASK_MASTER_SHEET in xl.sheet_names:
        df, hit = pd.read_excel(task_list_path, sheet_name=TASK_MASTER_SHEET), TASK_MASTER_SHEET
    else:
        for sh in xl.sheet_names:
            tmp = pd.read_excel(task_list_path, sheet_name=sh)
            if {'분析번호', '분析명', '실행여부'}.issubset(tmp.columns):
                df, hit = tmp, sh
                break

    if df is None:
        raise ValueError(
            f"'분析목록' 시트 또는 [분析번호/분析명/실행여부] 컬럼 포함 시트를 찾을 수 없습니다.\n"
            f"파일    : {task_list_path}\n"
            f"시트 목록: {xl.sheet_names}"
        )

    print(f"  [태스크 리스트] 시트='{hit}', 전체 {len(df)}개 항목")

    flag   = df['실행여부'].astype(str).str.strip().str.upper()
    active = df[flag.isin(['Y', 'O'])][['분析번호', '분析명']].dropna(subset=['분析번호'])
    tasks  = [(int(row['분析번호']), str(row['분析명']).strip()) for _, row in active.iterrows()]

    print(f"  [태스크 리스트] 실행 대상 {len(tasks)}개: {[f'{n}_{nm}' for n, nm in tasks]}")
    return tasks


def load_analysis_params(task_list_path: str, analysis_name: str) -> list:
    """
    분析명과 동일한 시트(또는 '분析명_파라미터' 시트)에서
    실행여부=Y/O 인 파라미터 행을 dict 리스트로 반환.

    파라미터 시트가 없으면 [{}] — 기본값으로 1회 실행.
    (공휴일전표, 중복전표 등 파라미터 불필요 분析에 해당)
    """
    xl = pd.ExcelFile(task_list_path)

    for candidate in [analysis_name, f'{analysis_name}_파라미터']:
        if candidate not in xl.sheet_names:
            continue

        df = pd.read_excel(task_list_path, sheet_name=candidate).dropna(how='all')
        if '실행여부' in df.columns:
            flag = df['실행여부'].astype(str).str.strip().str.upper()
            df   = df[flag.isin(['Y', 'O'])].copy()

        params_list = df.to_dict('records')
        print(f"    └ 파라미터 시트 '{candidate}': {len(params_list)}행 활성")
        return params_list

    return [{}]   # 파라미터 시트 없음 → 기본값 1회 실행


# ═══════════════════════════════════════════════════════════════════════════════
# 3. 분개장 데이터 로드
# ═══════════════════════════════════════════════════════════════════════════════

def load_journal(raw_data_dir: str) -> pd.DataFrame:
    """
    raw_data/ 폴더에서 분개장 파일 자동 탐색 후 로드.
    탐색 순서: 당기*분개장*.xlsx → *분개장*.xlsx → *journal*.xlsx
    """
    if not os.path.isdir(raw_data_dir):
        raise FileNotFoundError(f"raw_data 폴더가 없습니다: {raw_data_dir}")

    found = []
    for pat in ['당기*분개장*.xlsx', '*분개장*.xlsx', '*journal*.xlsx']:
        found = glob.glob(os.path.join(raw_data_dir, pat))
        if found:
            break

    if not found:
        raise FileNotFoundError(
            f"분개장 파일을 찾을 수 없습니다: {raw_data_dir}\n"
            f"파일명에 '분개장' 또는 'journal'이 포함되어야 합니다."
        )

    filepath = found[0]
    print(f"  [데이터 로드] {os.path.basename(filepath)}")

    df = pd.read_excel(filepath)
    df.columns = df.columns.str.strip()

    if '날짜' in df.columns:
        df['날짜'] = pd.to_datetime(df['날짜'], errors='coerce')

    print(f"  [데이터 로드] {len(df):,}건 로드 완료. 컬럼: {df.columns.tolist()}")
    return df


# ═══════════════════════════════════════════════════════════════════════════════
# 4. 분析 함수
#    모든 함수 시그니처: (df: DataFrame, params: dict = None) → DataFrame
# ═══════════════════════════════════════════════════════════════════════════════

def analyze_holiday_entries(df: pd.DataFrame, params: dict = None) -> pd.DataFrame:
    """1. 공휴일·주말 전표 — 파라미터 없이 실행."""
    if '날짜' not in df.columns:
        return pd.DataFrame([['날짜 컬럼 없음']], columns=['오류'])

    is_weekend = df['날짜'].dt.dayofweek >= 5

    try:
        import holidays as hd
        yr_range  = range(int(df['날짜'].dt.year.min()), int(df['날짜'].dt.year.max()) + 1)
        kr_hol    = hd.KoreaHolidays(years=yr_range)
        is_holiday = df['날짜'].dt.date.apply(lambda d: d in kr_hol)
    except ImportError:
        is_holiday = pd.Series(False, index=df.index)

    flagged        = df[is_weekend | is_holiday].copy()
    flagged['요일'] = flagged['날짜'].dt.day_name()
    return flagged.sort_values('날짜')


def analyze_duplicate_entries(df: pd.DataFrame, params: dict = None) -> pd.DataFrame:
    """2. 중복 전표 — 파라미터 없이 실행."""
    key_cols = [c for c in ['날짜', '계정과목', '차변', '대변', '거래처명'] if c in df.columns]
    dup_mask  = df.duplicated(subset=key_cols, keep=False)
    flagged   = df[dup_mask].copy()
    flagged['중복건수'] = flagged.groupby(key_cols)['날짜'].transform('count')
    return flagged.sort_values(key_cols)


def analyze_benford(df: pd.DataFrame, params: dict = None) -> pd.DataFrame:
    """
    3. 벤포드 법칙
    파라미터: 계정과목, 금액열
    |차이| > 5%p 이면 이상여부 'Y' 표시.
    """
    BENFORD = {1: 30.1, 2: 17.6, 3: 12.5, 4: 9.7,
               5: 7.9,  6: 6.7,  7: 5.8,  8: 5.1, 9: 4.6}

    params      = params or {}
    sub, a_cols = _filter_df(df, params)
    records     = []

    for col in a_cols:
        pos = sub[col].replace(0, pd.NA).dropna()
        pos = pos[pos > 0]
        if pos.empty:
            continue

        first_digits = pos.apply(lambda x: int(str(int(x))[0]))
        total        = len(first_digits)

        for d in range(1, 10):
            cnt  = (first_digits == d).sum()
            pct  = cnt / total * 100
            diff = pct - BENFORD[d]
            records.append({
                '금액열':          col,
                '첫째자리':        d,
                '실제건수':        int(cnt),
                '실제비율(%)':     round(pct, 2),
                '벤포드기댓값(%)':  BENFORD[d],
                '차이(%p)':        round(diff, 2),
                '이상여부':        'Y' if abs(diff) > 5 else '',
            })

    return pd.DataFrame(records)


def analyze_large_amounts(df: pd.DataFrame, params: dict = None) -> pd.DataFrame:
    """
    5. 거액 전표
    파라미터: 계정과목, 금액열, 기준백분위(%) (기본 99)
    """
    params      = params or {}
    pct         = _safe_float(params.get('기준백분위(%)'), default=99.0)
    sub, a_cols = _filter_df(df, params)
    records     = []

    for col in a_cols:
        series = sub[sub[col] > 0][col]
        if series.empty:
            continue
        cutoff  = series.quantile(pct / 100)
        flagged = sub[sub[col] >= cutoff].copy()
        flagged['금액열']   = col
        flagged['기준금액'] = cutoff
        records.append(flagged)

    return pd.concat(records, ignore_index=True) if records else pd.DataFrame()


def analyze_custom_search(df: pd.DataFrame, params: dict = None) -> pd.DataFrame:
    """
    6. 상세검색
    파라미터: 작업명, 계정과목, 거래처명, 적요, 금액유형, 최소금액, 최대금액
    금액유형 값: 차변만 / 대변만 / 양방향 (기본 양방향)
    """
    params   = params or {}
    acct     = _nv(params.get('계정과목', ''))
    vendor   = _nv(params.get('거래처명', ''))
    note     = _nv(params.get('적요', ''))
    amt_type = _nv(params.get('금액유형', ''), blank_vals=('nan', 'none', ''))
    min_amt  = params.get('최소금액')
    max_amt  = params.get('최대금액')

    mask = pd.Series(True, index=df.index)
    if acct   and '계정과목' in df.columns:
        mask &= df['계정과목'].astype(str).str.contains(acct, na=False)
    if vendor  and '거래처명' in df.columns:
        mask &= df['거래처명'].astype(str).str.contains(vendor, na=False)
    if note    and '적요' in df.columns:
        mask &= df['적요'].astype(str).str.contains(note, na=False)

    result = df[mask].copy()

    if amt_type == '차변만' and '차변' in result.columns:
        result = result[result['차변'] > 0]
    elif amt_type == '대변만' and '대변' in result.columns:
        result = result[result['대변'] > 0]

    amt_cols = [c for c in ('차변', '대변') if c in result.columns]
    if amt_cols:
        max_series = result[amt_cols].max(axis=1)
        if min_amt is not None:
            try:
                if not pd.isna(float(min_amt)):
                    result = result[max_series >= float(min_amt)]
            except (TypeError, ValueError):
                pass
        if max_amt is not None:
            try:
                if not pd.isna(float(max_amt)):
                    result = result[max_series <= float(max_amt)]
            except (TypeError, ValueError):
                pass

    return result.sort_values('날짜') if '날짜' in result.columns else result


def analyze_round_numbers(df: pd.DataFrame, params: dict = None) -> pd.DataFrame:
    """
    10. 라운드넘버
    파라미터: 계정과목, 금액열, 최소금액 (기본 100,000원)
    최소금액 이상인 단위 배수만 탐지.
    """
    params      = params or {}
    min_unit    = _safe_float(params.get('최소금액'), default=100_000.0)
    ALL_UNITS   = [100_000, 500_000, 1_000_000, 5_000_000, 10_000_000]
    units       = [u for u in ALL_UNITS if u >= min_unit]
    sub, a_cols = _filter_df(df, params)
    records     = []

    for col in a_cols:
        work = sub[sub[col] > 0].copy()
        work['라운드단위'] = work[col].apply(
            lambda x: next(
                (f'{u:,}원 배수' for u in sorted(units, reverse=True) if x % u == 0),
                None
            )
        )
        flagged = work[work['라운드단위'].notna()].assign(금액열=col)
        records.append(flagged)

    if not records:
        return pd.DataFrame()

    result   = pd.concat(records, ignore_index=True)
    amt_col  = next((c for c in ('차변', '대변', '금액') if c in result.columns), None)
    out_cols = [c for c in ['날짜', '전표번호', '계정과목', '금액열', amt_col,
                             '적요', '거래처명', '라운드단위'] if c and c in result.columns]
    return (result[out_cols].sort_values(amt_col, ascending=False)
            if amt_col else result[out_cols])


# ═══════════════════════════════════════════════════════════════════════════════
# 분析 레지스트리
# 새 분析 추가 시: 함수 작성 후 이곳에 {번호: (名, 함수)} 한 줄만 추가
# ═══════════════════════════════════════════════════════════════════════════════

ANALYSIS_REGISTRY: dict = {
    1:  ('공휴일전표', analyze_holiday_entries),
    2:  ('중복전표',   analyze_duplicate_entries),
    3:  ('벤포드분析', analyze_benford),
    5:  ('거액전표',   analyze_large_amounts),
    6:  ('상세검색',   analyze_custom_search),
    10: ('라운드넘버', analyze_round_numbers),
}


# ═══════════════════════════════════════════════════════════════════════════════
# 5. 결과 통합 저장
# ═══════════════════════════════════════════════════════════════════════════════

def save_results(results: dict, output_dir: str) -> str:
    """
    { '3_벤포드분析_외상매출금_차변': df, ... } →
    {output_dir}/분析결과_temp.xlsx 에 시트별 저장.
    추후 injector.py 가 이 파일을 읽어 조서에 주입한다.
    """
    os.makedirs(output_dir, exist_ok=True)
    out_path = os.path.join(output_dir, '분析결과_temp.xlsx')

    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        for sheet, df in results.items():
            safe = sheet[:31]
            if df is None or df.empty:
                pd.DataFrame([['결과 없음']]).to_excel(
                    writer, sheet_name=safe, index=False, header=False
                )
            else:
                df.to_excel(writer, sheet_name=safe, index=False)

    return out_path


# ═══════════════════════════════════════════════════════════════════════════════
# 메인
# ═══════════════════════════════════════════════════════════════════════════════

def main():
    parser = argparse.ArgumentParser(description='분개장 분析 자동화')
    parser.add_argument('company', nargs='?', help='고객사 이름 (예: graphy)')
    args = parser.parse_args()

    company_name = (args.company or input('고객사 이름을 입력하세요: ')).strip()
    if not company_name:
        print('[오류] 고객사 이름이 비어 있습니다.')
        sys.exit(1)

    print(f'\n{"=" * 60}')
    print(f'  분개장 분析 자동화 — {company_name}')
    print(f'{"=" * 60}')

    # 1) 경로 세팅
    paths = resolve_paths(company_name)
    print('\n[경로 세팅]')
    for k, v in paths.items():
        print(f'  {k:<14}: {v}')

    # 2) 태스크 리스트 로드
    print('\n[태스크 리스트]')
    try:
        active_tasks = load_active_tasks(paths['task_list'])
    except (FileNotFoundError, ValueError) as e:
        print(f'[오류] {e}')
        sys.exit(1)

    if not active_tasks:
        print('  실행할 분析이 없습니다 (Y/O 항목 없음).')
        sys.exit(0)

    # 3) 분개장 로드
    print('\n[분개장 데이터 로드]')
    try:
        journal_df = load_journal(paths['raw_data'])
    except FileNotFoundError as e:
        print(f'[오류] {e}')
        sys.exit(1)

    # 4) 분析 순차 실행
    print('\n[분析 실행]')
    results: dict = {}

    for task_no, task_name in active_tasks:
        if task_no not in ANALYSIS_REGISTRY:
            print(f'  [{task_no:>3}] {task_name:<20} → [건너뜀] 등록된 함수 없음')
            continue

        _, func = ANALYSIS_REGISTRY[task_no]
        print(f'  [{task_no:>3}] {task_name}')
        params_list = load_analysis_params(paths['task_list'], task_name)

        for params in params_list:
            key = _result_key(task_no, task_name, params)
            try:
                result_df   = func(journal_df, params)
                results[key] = result_df
                n = len(result_df) if result_df is not None else 0
                print(f'    → [{key}] {n:,}건')
            except Exception as e:
                results[key] = pd.DataFrame([['오류: ' + str(e)]])
                print(f'    → [{key}] [오류] {e}')

    # 5) 결과 저장
    if results:
        print('\n[결과 저장]')
        out_path = save_results(results, paths['output'])
        print(f'  파일  : {out_path}')
        print(f'  시트  : {list(results.keys())}')

    print(f'\n{"=" * 60}')
    print(f'  완료 — {len(results)}개 분析 결과 저장')
    print(f'{"=" * 60}\n')


if __name__ == '__main__':
    main()
