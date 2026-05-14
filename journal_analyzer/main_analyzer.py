#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
journal_analyzer/main_analyzer.py
분개장 분석 자동화 메인 스크립트

실행 예시:
    python main_analyzer.py graphy
    python main_analyzer.py          # → 입력창 표시
"""

import sys
import os
import argparse
import glob
import warnings

import pandas as pd

warnings.filterwarnings('ignore', category=pd.errors.DtypeWarning)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# task_list_{company}.xlsx 에서 분析 목록을 담은 시트명
TASK_SHEET_NAME = '분析목록'


# ═══════════════════════════════════════════════════════════════════════════════
# 1. 경로 세팅
# ═══════════════════════════════════════════════════════════════════════════════

def resolve_paths(company_name: str) -> dict:
    """고객사 이름을 기준으로 하위 경로를 동적으로 세팅."""
    company_dir = os.path.normpath(os.path.join(BASE_DIR, '..', company_name))
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
    task_list_{company}.xlsx 의 '분析번호 / 분析명 / 실행여부' 컬럼에서
    실행여부가 Y 또는 O인 항목을 (번호: int, 이름: str) 튜플 리스트로 반환.

    시트 탐색 순서:
      1) TASK_SHEET_NAME('분析목록')으로 직접 접근
      2) 없으면 전체 시트를 순회해 해당 컬럼을 포함한 시트를 자동 탐색
    """
    if not os.path.isfile(task_list_path):
        raise FileNotFoundError(f"task_list 파일이 없습니다: {task_list_path}")

    xl  = pd.ExcelFile(task_list_path)
    df  = None
    hit = None

    if TASK_SHEET_NAME in xl.sheet_names:
        df  = pd.read_excel(task_list_path, sheet_name=TASK_SHEET_NAME)
        hit = TASK_SHEET_NAME
    else:
        for sh in xl.sheet_names:
            tmp = pd.read_excel(task_list_path, sheet_name=sh)
            if {'분析번호', '분析명', '실행여부'}.issubset(tmp.columns):
                df, hit = tmp, sh
                break

    if df is None:
        raise ValueError(
            f"'분析번호', '분析명', '실행여부' 컬럼을 가진 시트를 찾을 수 없습니다.\n"
            f"파일    : {task_list_path}\n"
            f"시트 목록: {xl.sheet_names}\n"
            f"힌트    : '{TASK_SHEET_NAME}' 시트를 추가하거나 컬럼명을 확인하세요."
        )

    print(f"  [태스크 리스트] 시트='{hit}', 전체 {len(df)}개 항목")

    flag = df['실행여부'].astype(str).str.strip().str.upper()
    active = df[flag.isin(['Y', 'O'])][['분析번호', '분析명']].dropna(subset=['분析번호'])

    tasks = [(int(row['분析번호']), str(row['분析명']).strip()) for _, row in active.iterrows()]
    print(f"  [태스크 리스트] 실행 대상 {len(tasks)}개: {[f'{n}_{nm}' for n, nm in tasks]}")
    return tasks


# ═══════════════════════════════════════════════════════════════════════════════
# 3. 분개장 데이터 로드
# ═══════════════════════════════════════════════════════════════════════════════

def load_journal(raw_data_dir: str) -> pd.DataFrame:
    """
    raw_data 폴더에서 분개장 파일을 자동 탐색 후 로드.
    탐색 패턴: 당기*분개장*.xlsx → *분개장*.xlsx → *journal*.xlsx
    (DtypeWarning 은 모듈 상단 warnings.filterwarnings 로 억제)
    """
    if not os.path.isdir(raw_data_dir):
        raise FileNotFoundError(f"raw_data 폴더가 없습니다: {raw_data_dir}")

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
# ═══════════════════════════════════════════════════════════════════════════════

def analyze_holiday_entries(df: pd.DataFrame) -> pd.DataFrame:
    """1. 공휴일·주말 전표 — 토·일·공휴일에 입력된 전표 추출."""
    if '날짜' not in df.columns:
        return pd.DataFrame([['날짜 컬럼 없음']], columns=['오류'])

    is_weekend = df['날짜'].dt.dayofweek >= 5

    try:
        import holidays as hd
        yr_min = int(df['날짜'].dt.year.min())
        yr_max = int(df['날짜'].dt.year.max())
        kr_hol = hd.KoreaHolidays(years=range(yr_min, yr_max + 1))
        is_holiday = df['날짜'].dt.date.apply(lambda d: d in kr_hol)
    except ImportError:
        is_holiday = pd.Series(False, index=df.index)

    flagged = df[is_weekend | is_holiday].copy()
    flagged['요일'] = flagged['날짜'].dt.day_name()
    return flagged.sort_values('날짜')


def analyze_duplicate_entries(df: pd.DataFrame) -> pd.DataFrame:
    """2. 중복 전표 — 날짜·계정과목·금액·거래처 조합이 2회 이상인 전표 추출."""
    key_cols = [c for c in ['날짜', '계정과목', '차변', '대변', '거래처명'] if c in df.columns]
    dup_mask  = df.duplicated(subset=key_cols, keep=False)
    flagged   = df[dup_mask].copy()
    flagged['중복건수'] = flagged.groupby(key_cols)['날짜'].transform('count')
    return flagged.sort_values(key_cols)


def analyze_benford(df: pd.DataFrame) -> pd.DataFrame:
    """
    3. 벤포드 법칙 분析
    차·대변 금액의 첫째 유효숫자(1~9) 실제 비율과 벤포드 기댓값을 비교.
    |차이| > 5%p 인 경우 이상여부 'Y' 표시.
    """
    BENFORD = {1: 30.1, 2: 17.6, 3: 12.5, 4: 9.7, 5: 7.9, 6: 6.7, 7: 5.8, 8: 5.1, 9: 4.6}
    amount_cols = [c for c in df.columns if c in ('차변', '대변', '금액')]
    records = []

    for col in amount_cols:
        pos = df[col].replace(0, pd.NA).dropna()
        pos = pos[pos > 0]
        if pos.empty:
            continue

        first_digits = pos.apply(lambda x: int(str(int(x))[0]))
        total = len(first_digits)

        for d in range(1, 10):
            cnt  = (first_digits == d).sum()
            pct  = cnt / total * 100
            diff = pct - BENFORD[d]
            records.append({
                '금액열':         col,
                '첫째자리':       d,
                '실제건수':       cnt,
                '실제비율(%)':    round(pct, 2),
                '벤포드기댓값(%)': BENFORD[d],
                '차이(%p)':       round(diff, 2),
                '이상여부':       'Y' if abs(diff) > 5 else '',
            })

    return pd.DataFrame(records)


def analyze_large_amounts(df: pd.DataFrame, percentile: float = 99.0) -> pd.DataFrame:
    """5. 거액 전표 — 차·대변 금액 상위 1% 이상인 전표 추출."""
    records = []
    for col in [c for c in df.columns if c in ('차변', '대변')]:
        sub = df[df[col] > 0].copy()
        if sub.empty:
            continue
        cutoff = sub[col].quantile(percentile / 100)
        flagged = sub[sub[col] >= cutoff].copy()
        flagged['금액열']   = col
        flagged['기준금액'] = cutoff
        records.append(flagged)

    return pd.concat(records, ignore_index=True) if records else pd.DataFrame()


def analyze_round_numbers(df: pd.DataFrame) -> pd.DataFrame:
    """
    10. 라운드넘버 분析
    10만 / 50만 / 100만 / 500만 / 1000만 단위 배수 금액 전표 탐지.
    """
    THRESHOLDS = [100_000, 500_000, 1_000_000, 5_000_000, 10_000_000]
    records = []

    for col in [c for c in df.columns if c in ('차변', '대변', '금액')]:
        sub = df[df[col] > 0].copy()
        sub['라운드단위'] = sub[col].apply(
            lambda x: next(
                (f'{t:,}원 배수' for t in sorted(THRESHOLDS, reverse=True) if x % t == 0),
                None
            )
        )
        flagged = sub[sub['라운드단위'].notna()].copy()
        flagged['금액열'] = col
        records.append(flagged)

    if not records:
        return pd.DataFrame()

    result   = pd.concat(records, ignore_index=True)
    out_cols = [c for c in ['날짜', '전표번호', '계정과목', '금액열', col,
                             '적요', '거래처명', '라운드단위'] if c in result.columns]
    return result[out_cols].sort_values(col if col in result.columns else out_cols[0],
                                        ascending=False)


# ═══════════════════════════════════════════════════════════════════════════════
# 분析 레지스트리 — { 분析번호: (시트명, 함수) }
# 새 분析 모듈을 추가할 때 여기에만 등록하면 자동으로 실행 대상에 포함됨
# ═══════════════════════════════════════════════════════════════════════════════

ANALYSIS_REGISTRY: dict = {
    1:  ('공휴일전표',   analyze_holiday_entries),
    2:  ('중복전표',     analyze_duplicate_entries),
    3:  ('벤포드분析',   analyze_benford),
    5:  ('거액전표',     analyze_large_amounts),
    10: ('라운드넘버',   analyze_round_numbers),
}


# ═══════════════════════════════════════════════════════════════════════════════
# 5. 결과 통합 저장
# ═══════════════════════════════════════════════════════════════════════════════

def save_results(results: dict, output_dir: str) -> str:
    """
    { '3_벤포드분析': df, '10_라운드넘버': df, ... } 를
    {company}/results/분析결과_temp.xlsx 에 시트별로 저장.
    추후 injector.py 가 이 파일을 읽어 조서에 주입한다.
    """
    os.makedirs(output_dir, exist_ok=True)
    out_path = os.path.join(output_dir, '분析결과_temp.xlsx')

    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        for sheet_name, df in results.items():
            safe = sheet_name[:31]                        # Excel 시트명 최대 31자
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
    print(f'  분개장 분析 자동화 시작 — {company_name}')
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
        print('  실행할 분析이 없습니다 (실행여부=Y/O 항목 없음). 종료합니다.')
        sys.exit(0)

    # 3) 분개장 데이터 로드
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
            print(f'  [{task_no:>3}] {task_name:<22} → [건너뜀] 등록된 함수 없음')
            continue

        _, func = ANALYSIS_REGISTRY[task_no]
        try:
            result_df = func(journal_df)
            key = f'{task_no}_{task_name}'
            results[key] = result_df
            n = len(result_df) if result_df is not None else 0
            print(f'  [{task_no:>3}] {task_name:<22} → {n:,}건')
        except Exception as e:
            key = f'{task_no}_{task_name}'
            results[key] = pd.DataFrame([['오류: ' + str(e)]])
            print(f'  [{task_no:>3}] {task_name:<22} → [오류] {e}')

    # 5) 결과 저장
    if results:
        print('\n[결과 저장]')
        out_path = save_results(results, paths['output'])
        print(f'  파일   : {out_path}')
        print(f'  시트   : {list(results.keys())}')

    print(f'\n{"=" * 60}')
    print(f'  완료 — {len(results)}개 분析 결과 저장')
    print(f'{"=" * 60}\n')


if __name__ == '__main__':
    main()
