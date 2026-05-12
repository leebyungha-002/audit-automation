"""
K-IFRS 1116 리스 완전성(Completeness) 검토 스크립트
- input_data/ 폴더의 기중 비용 원장 엑셀을 읽어
- 리스 관련 키워드가 적요에 포함된 거래를 추출·집계하고
- output/리스완전성검토_후보목록.xlsx 로 저장한다.
"""

import glob
import os
import re

import pandas as pd
from openpyxl.utils import get_column_letter

# ── 리스 탐지 키워드 ──────────────────────────────────────────────────────────
# 적요에 아래 단어 중 하나라도 포함된 거래를 리스 후보로 추출
LEASE_KEYWORDS = [
    '호', '빌딩', '건물', '임차', '차량', '렌탈', '렌트', '리스',
    '사무실', '복합기', '정수기', '냉난방', '에어컨', '주차', '창고',
]

# ── 경로 설정 ─────────────────────────────────────────────────────────────────
BASE_DIR   = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR  = os.path.join(BASE_DIR, 'input_data')
OUTPUT_DIR = os.path.join(BASE_DIR, 'output')
OUTPUT_FILE = os.path.join(OUTPUT_DIR, '리스완전성검토_후보목록.xlsx')

KEYWORD_PATTERN = '|'.join(re.escape(kw) for kw in LEASE_KEYWORDS)

REQUIRED_COLS = {'일자', '계정과목', '적요', '거래처', '차변'}


# ── 1. 데이터 로드 ─────────────────────────────────────────────────────────────
def load_ledger() -> pd.DataFrame:
    files = glob.glob(os.path.join(INPUT_DIR, '*.xlsx'))
    files += glob.glob(os.path.join(INPUT_DIR, '*.xls'))
    if not files:
        raise FileNotFoundError(
            f"[오류] input_data 폴더에 엑셀 파일이 없습니다.\n  경로: {INPUT_DIR}"
        )

    frames = []
    for f in files:
        df = pd.read_excel(f)
        missing = REQUIRED_COLS - set(df.columns)
        if missing:
            raise ValueError(
                f"[오류] 필수 컬럼 누락 — {os.path.basename(f)}: {missing}"
            )
        frames.append(df)
        print(f"  로드: {os.path.basename(f)}  ({len(df):,}행)")

    return pd.concat(frames, ignore_index=True)


# ── 2. 결측치 처리 + 키워드 필터링 ────────────────────────────────────────────
def filter_lease_rows(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df['적요']  = df['적요'].fillna('').astype(str).str.strip()
    df['거래처'] = df['거래처'].fillna('(미기재)').astype(str).str.strip()
    df['차변']  = pd.to_numeric(df['차변'], errors='coerce').fillna(0)

    mask = df['적요'].str.contains(
        KEYWORD_PATTERN, flags=re.IGNORECASE, regex=True, na=False
    )
    return df[mask].copy()


# ── 3. 거래처 × 계정과목 집계 ─────────────────────────────────────────────────
def _top_remarks(series: pd.Series, n: int = 2) -> str:
    """가장 빈번한 적요 상위 n개를 ' / ' 로 결합하여 반환."""
    cleaned = series[series.str.strip() != '']
    if cleaned.empty:
        return ''
    top = cleaned.value_counts().head(n).index.tolist()
    return ' / '.join(top)


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
    # 감사인 수기 체크용 빈 컬럼
    result['리스인식여부(O/X)'] = ''
    result['비고'] = ''
    return result


# ── 4. 엑셀 저장 (천단위 콤마 서식 적용) ──────────────────────────────────────
def save_excel(df: pd.DataFrame) -> None:
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    col_labels = {
        '연간_총발생액': '연간 총 발생액',
        '거래_발생건수': '거래 발생건수',
        '대표_적요':    '대표 적요',
    }
    df = df.rename(columns=col_labels)

    with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='리스후보목록')
        ws = writer.sheets['리스후보목록']

        # 천단위 콤마 서식: '연간 총 발생액' 컬럼
        amount_col_idx = df.columns.get_loc('연간 총 발생액') + 1  # openpyxl 1-based
        for row in ws.iter_rows(
            min_row=2, max_row=ws.max_row,
            min_col=amount_col_idx, max_col=amount_col_idx
        ):
            for cell in row:
                cell.number_format = '#,##0'

        # 헤더 굵게
        from openpyxl.styles import Font, PatternFill, Alignment
        header_fill = PatternFill('solid', fgColor='D9E1F2')
        for cell in ws[1]:
            cell.font = Font(bold=True)
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center')

        # 컬럼 너비 자동 조정
        for col_idx, col_name in enumerate(df.columns, start=1):
            col_letter = get_column_letter(col_idx)
            header_len = len(str(col_name))
            if len(df) > 0:
                max_data_len = df.iloc[:, col_idx - 1].astype(str).str.len().max()
                max_data_len = 0 if pd.isna(max_data_len) else int(max_data_len)
            else:
                max_data_len = 0
            ws.column_dimensions[col_letter].width = min(
                max(header_len, max_data_len) + 4, 60
            )

        # 행 높이
        ws.row_dimensions[1].height = 18

    print(f"\n저장 완료: {OUTPUT_FILE}")
    print(f"  후보 건수(거래처×계정과목): {len(df):,}건")
    print(f"  연간 총 발생액 합계: {df['연간 총 발생액'].sum():,.0f}원")


# ── main ──────────────────────────────────────────────────────────────────────
def main() -> None:
    print("=" * 55)
    print("  K-IFRS 1116 리스 완전성 검토 — 후보 추출")
    print("=" * 55)

    print("\n[1/3] 원장 데이터 로드")
    raw = load_ledger()
    print(f"  전체 행 수: {len(raw):,}")

    print("\n[2/3] 리스 키워드 필터링")
    print(f"  적용 키워드: {LEASE_KEYWORDS}")
    filtered = filter_lease_rows(raw)
    print(f"  필터링 후 행 수: {len(filtered):,}  "
          f"(전체 대비 {len(filtered)/len(raw)*100:.1f}%)")

    if filtered.empty:
        print("\n  [안내] 리스 키워드에 해당하는 거래 내역이 없습니다.")
        return

    print("\n[3/3] 거래처·계정과목별 집계 및 저장")
    result = aggregate(filtered)
    save_excel(result)

    print("\n완료. output 폴더의 결과 파일에서 '리스인식여부(O/X)' 컬럼을 체크해 주세요.")
    print("=" * 55)


if __name__ == '__main__':
    main()
