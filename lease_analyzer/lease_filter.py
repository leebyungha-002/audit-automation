"""
K-IFRS 1116 리스 완전성(Completeness) 검토 스크립트

실행 방법:
  [모드 1] 기중비용원장 직접 분석 (input_data/ 폴더 사용)
      python lease_filter.py

  [모드 2] Playwright 상세검색 결과 파일 연계
      python lease_filter.py --company dae_il
      python lease_filter.py --company dae_il --no-filter
"""

import argparse
import glob
import io
import os
import re
import sys

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ── 리스 탐지 키워드 ──────────────────────────────────────────────────────────
LEASE_KEYWORDS = [
    '호', '빌딩', '건물', '임차', '차량', '렌탈', '렌트', '리스',
    '사무실', '복합기', '정수기', '냉난방', '에어컨', '주차', '창고',
]

# ── 경로 설정 ─────────────────────────────────────────────────────────────────
BASE_DIR    = os.path.dirname(os.path.abspath(__file__))
PROJECT_ROOT = os.path.dirname(BASE_DIR)   # audit-automation/
INPUT_DIR   = os.path.join(BASE_DIR, 'input_data')
OUTPUT_DIR  = os.path.join(BASE_DIR, 'output')

KEYWORD_PATTERN = '|'.join(re.escape(kw) for kw in LEASE_KEYWORDS)
REQUIRED_COLS   = {'일자', '계정과목', '적요', '거래처', '차변'}

# 계정과목 코드 정규화 패턴: '205_이자비용(93100)' → '이자비용'
_ACCOUNT_RE = re.compile(r'^\d+_(.+?)\s*\(\d+\)\s*$')


# ── 공통 유틸 ─────────────────────────────────────────────────────────────────
def _clean_account(name: str) -> str:
    m = _ACCOUNT_RE.match(str(name).strip())
    return m.group(1).strip() if m else str(name).strip()


def _top_remarks(series: pd.Series, n: int = 2) -> str:
    cleaned = series[series.str.strip() != '']
    if cleaned.empty:
        return ''
    return ' / '.join(cleaned.value_counts().head(n).index.tolist())


# ── 모드 1: input_data/ 원장 로드 ────────────────────────────────────────────
def load_ledger() -> pd.DataFrame:
    files = glob.glob(os.path.join(INPUT_DIR, '*.xlsx'))
    files += glob.glob(os.path.join(INPUT_DIR, '*.xls'))
    if not files:
        raise FileNotFoundError(
            f"[오류] input_data 폴더에 엑셀 파일이 없습니다.\n  경로: {INPUT_DIR}"
        )
    frames = []
    for f in sorted(files):
        df = pd.read_excel(f)
        missing = REQUIRED_COLS - set(df.columns)
        if missing:
            raise ValueError(f"[오류] 필수 컬럼 누락 - {os.path.basename(f)}: {missing}")
        frames.append(df)
        print(f"  로드: {os.path.basename(f)}  ({len(df):,}행)")
    return pd.concat(frames, ignore_index=True)


# ── 모드 2: 회사 results/ 상세검색 결과 로드 ──────────────────────────────────
def _normalize_result_sheet(df: pd.DataFrame, sheet_name: str) -> pd.DataFrame:
    """상세검색 결과 컬럼 → 표준 컬럼 정규화."""
    df = df.copy()

    # 날짜 → 일자
    if '날짜' in df.columns:
        df = df.rename(columns={'날짜': '일자'})

    # 적요란 → 적요 (실제 설명은 적요란에 있음)
    if '적요란' in df.columns:
        raw_desc = df['적요란'].fillna('')
        df['적요'] = raw_desc.astype(str).str.strip()

    # 계정과목 코드 정규화
    if '계정과목' in df.columns:
        df['계정과목'] = df['계정과목'].astype(str).apply(_clean_account)

    # 계정과목이 모두 비어있으면 시트명으로 보완
    if '계정과목' not in df.columns or df['계정과목'].replace('', pd.NA).isna().all():
        df['계정과목'] = sheet_name

    # 전기이월 행 제거 (집계 왜곡 방지)
    if '적요' in df.columns:
        df = df[~df['적요'].str.contains('전기이월', na=False)]

    return df


def load_from_result(company: str) -> pd.DataFrame:
    """회사 results/ 폴더에서 리스 완전성 결과 파일을 찾아 전체 시트 병합."""
    results_dir = os.path.join(PROJECT_ROOT, company, 'results')
    if not os.path.isdir(results_dir):
        raise FileNotFoundError(f"[오류] results 폴더가 없습니다: {results_dir}")

    candidates = []
    for pat in ['*리스*완전성*.xlsx', '*리스완전성*.xlsx', '*리스거래*.xlsx']:
        candidates += glob.glob(os.path.join(results_dir, pat))
    candidates = sorted(set(candidates))   # 최신순 정렬 (파일명에 날짜 포함)

    if not candidates:
        raise FileNotFoundError(
            f"[오류] 리스 완전성 결과 파일을 찾을 수 없습니다.\n"
            f"  탐색 경로: {results_dir}\n"
            f"  Playwright 상세검색 시나리오를 먼저 실행해 주세요."
        )

    target = candidates[-1]   # 가장 최신 파일
    print(f"  대상 파일: {os.path.basename(target)}")

    xl = pd.ExcelFile(target)
    print(f"  시트 목록: {xl.sheet_names}")

    frames = []
    for sheet in xl.sheet_names:
        raw = pd.read_excel(target, sheet_name=sheet)
        normalized = _normalize_result_sheet(raw, sheet)
        frames.append(normalized)
        print(f"  시트 '{sheet}': {len(normalized):,}행")

    merged = pd.concat(frames, ignore_index=True)

    # 필수 컬럼 누락 체크
    missing = REQUIRED_COLS - set(merged.columns)
    if missing:
        raise ValueError(
            f"[오류] 정규화 후에도 필수 컬럼 누락: {missing}\n"
            f"  실제 컬럼: {list(merged.columns)}"
        )
    return merged


# ── 공통 파이프라인 ───────────────────────────────────────────────────────────
def preprocess(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df['적요']   = df['적요'].fillna('').astype(str).str.strip()
    df['거래처']  = df['거래처'].fillna('(미기재)').astype(str).str.strip()
    df['차변']   = pd.to_numeric(df['차변'], errors='coerce').fillna(0)
    return df


def filter_lease_rows(df: pd.DataFrame) -> pd.DataFrame:
    return df[df['적요'].str.contains(
        KEYWORD_PATTERN, flags=re.IGNORECASE, regex=True, na=False
    )].copy()


def aggregate(df: pd.DataFrame) -> pd.DataFrame:
    result = (
        df.groupby(['거래처', '계정과목'], as_index=False, sort=False)
        .agg(
            연간_총발생액=('차변', 'sum'),
            거래_발생건수=('차변', 'count'),
            대표_적요=('적요', _top_remarks),
        )
        .sort_values('연간_총발생액', ascending=False)
        .reset_index(drop=True)
    )
    result['리스인식여부(O/X)'] = ''
    result['비고'] = ''
    return result


def save_excel(df: pd.DataFrame, output_path: str) -> None:
    os.makedirs(os.path.dirname(output_path), exist_ok=True)

    col_labels = {
        '연간_총발생액': '연간 총 발생액',
        '거래_발생건수': '거래 발생건수',
        '대표_적요':    '대표 적요',
    }
    df = df.rename(columns=col_labels)

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='리스후보목록')
        ws = writer.sheets['리스후보목록']

        # 천단위 콤마 서식
        amt_col = df.columns.get_loc('연간 총 발생액') + 1
        for row in ws.iter_rows(min_row=2, max_row=ws.max_row,
                                min_col=amt_col, max_col=amt_col):
            for cell in row:
                cell.number_format = '#,##0'

        # 헤더 스타일
        hdr_fill = PatternFill('solid', fgColor='D9E1F2')
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.fill = hdr_fill
            cell.alignment = Alignment(horizontal='center', vertical='center')
        ws.row_dimensions[1].height = 18

        # 컬럼 너비 자동 조정
        for col_idx, col_name in enumerate(df.columns, start=1):
            header_len = len(str(col_name))
            data_len = (
                df.iloc[:, col_idx - 1].astype(str).str.len().max()
                if len(df) > 0 else 0
            )
            data_len = 0 if pd.isna(data_len) else int(data_len)
            ws.column_dimensions[get_column_letter(col_idx)].width = min(
                max(header_len, data_len) + 4, 60
            )

    print(f"\n저장 완료: {output_path}")
    print(f"  후보 건수(거래처x계정과목): {len(df):,}건")
    print(f"  연간 총 발생액 합계: {df['연간 총 발생액'].sum():,.0f}원")


# ── CLI ───────────────────────────────────────────────────────────────────────
def parse_args() -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description='K-IFRS 1116 리스 완전성 검토',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            '예시:\n'
            '  python lease_filter.py                     # input_data/ 원장 분석\n'
            '  python lease_filter.py --company dae_il    # Playwright 결과 연계\n'
            '  python lease_filter.py --company dae_il --no-filter  # 키워드 필터 생략\n'
        )
    )
    p.add_argument('--company', '-c', metavar='COMPANY',
                   help='회사 폴더명 (예: dae_il). 지정 시 results/*리스*완전성*.xlsx 자동 탐색')
    p.add_argument('--no-filter', action='store_true',
                   help='키워드 필터링 생략 (계정과목이 이미 리스 특정 계정인 경우)')
    return p.parse_args()


# ── main ──────────────────────────────────────────────────────────────────────
def main() -> None:
    args = parse_args()

    print('=' * 60)
    print('  K-IFRS 1116 리스 완전성 검토 - 후보 추출')
    if args.company:
        print(f'  모드: 상세검색 결과 연계  (회사: {args.company})')
    else:
        print('  모드: 기중비용원장 직접 분석')
    print('=' * 60)

    # ── 1. 데이터 로드
    print('\n[1/3] 데이터 로드')
    if args.company:
        raw = load_from_result(args.company)
        output_file = os.path.join(OUTPUT_DIR, f'{args.company}_리스완전성검토_후보목록.xlsx')
    else:
        raw = load_ledger()
        output_file = os.path.join(OUTPUT_DIR, '리스완전성검토_후보목록.xlsx')
    print(f'  전체 행 수: {len(raw):,}')

    # ── 2. 전처리 + 필터링
    print('\n[2/3] 전처리 및 필터링')
    df = preprocess(raw)

    if args.no_filter:
        filtered = df[df['차변'] > 0].copy()   # 차변 발생 건만
        print(f'  키워드 필터 생략 (--no-filter) → 차변 발생 건: {len(filtered):,}행')
    else:
        print(f'  적용 키워드: {LEASE_KEYWORDS}')
        filtered = filter_lease_rows(df)
        ratio = len(filtered) / len(df) * 100 if len(df) else 0
        print(f'  필터링 후: {len(filtered):,}행  (전체 대비 {ratio:.1f}%)')

    if filtered.empty:
        print('\n  [안내] 해당하는 거래 내역이 없습니다.')
        return

    # ── 3. 집계 + 저장
    print('\n[3/3] 거래처·계정과목별 집계 및 저장')
    result = aggregate(filtered)
    save_excel(result, output_file)

    print("\n완료. 결과 파일에서 '리스인식여부(O/X)' 컬럼을 체크해 주세요.")
    print('=' * 60)


if __name__ == '__main__':
    main()
