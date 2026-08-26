"""
이자비용 적정성 분석 (Reasonableness Test)
실행: python interest_expense_analysis.py <회사명>
예시: python interest_expense_analysis.py dae_il
입력: interest_analyzer/<회사>/input/이자비용적정성.xlsx
출력: interest_analyzer/<회사>/output/이자비용분석결과.xlsx
"""

import sys
from pathlib import Path

import numpy as np
import pandas as pd

# ── 이자율 설정 (시트명: 연이자율) ───────────────────────────────────────────
# 키: 엑셀 시트명과 정확히 일치해야 함
# 값: 연이자율 소수 (예: 3.62% → 0.0362)
INTEREST_RATES = {
    '하나 장기운전자금4(29320)':                          0.0362,
    '하나_장기운전자금-34142 _ 하나 장기운전자금-341':    0.0336,
}

PERIOD_START   = pd.Timestamp('2026-01-01')  # 당기 기초일
PERIOD_END     = pd.Timestamp('2026-12-31')  # 당기 기말일
INTEREST_SHEET = '이자비용'                   # 실제 이자비용 시트명


# ─────────────────────────────────────────────────────────────────────────────

def load_loan_sheet(df: pd.DataFrame) -> pd.DataFrame:
    """차입금 시트 전처리: 결측치 처리 → 전기이월 날짜 보정 → 정렬."""
    df = df.copy()
    df['차변'] = pd.to_numeric(df['차변'], errors='coerce').fillna(0)
    df['대변'] = pd.to_numeric(df['대변'], errors='coerce').fillna(0)

    # 전기이월 여부: 적요란 == '전기이월' 이거나 날짜가 비어있는 행
    df['_is_bfwd'] = (
        df['적요란'].astype(str).str.strip().eq('전기이월') | df['날짜'].isna()
    )
    # datetime 변환 먼저 → 이후 Timestamp 대입 시 dtype 불일치 경고 방지
    df['날짜'] = pd.to_datetime(df['날짜'], errors='coerce')
    df.loc[df['_is_bfwd'], '날짜'] = PERIOD_START
    df = df.dropna(subset=['날짜']).sort_values('날짜').reset_index(drop=True)
    return df


def calc_daily_balance(df: pd.DataFrame) -> pd.Series:
    """거래 데이터를 당기 전체 일별 잔액 Series로 변환."""
    df['잔액'] = (df['대변'] - df['차변']).cumsum()
    # 같은 날짜에 거래가 여러 건이면 마지막(당일 종가) 잔액 사용
    balance_by_day = df.groupby('날짜')['잔액'].last()
    date_range = pd.date_range(PERIOD_START, PERIOD_END, freq='D')
    return balance_by_day.reindex(date_range).ffill().fillna(0)


def calc_expected_interest(balance_daily: pd.Series, rate: float) -> float:
    """일별 잔액 × (연이자율 / 365) 합산 → 기대이자비용."""
    return float((balance_daily * rate / 365).sum())


def main():
    # ── 회사명 인자 확인 ──────────────────────────────────────────────────────
    root = Path(__file__).parent  # interest_analyzer/
    if len(sys.argv) < 2:
        print('[오류] 회사명을 인자로 지정하세요.')
        print('  예시: python interest_expense_analysis.py dae_il')
        # 회사 폴더 목록 안내
        companies = [p.name for p in root.iterdir() if p.is_dir() and not p.name.startswith(('_', '.'))]
        if companies:
            print(f'  사용 가능 회사: {", ".join(sorted(companies))}')
        sys.exit(1)

    company = sys.argv[1]
    company_dir = root / company
    input_path  = company_dir / 'input'  / '이자비용적정성.xlsx'
    output_dir  = company_dir / 'output'
    output_path = output_dir  / '이자비용분석결과.xlsx'

    if not input_path.exists():
        print(f'[오류] 파일 없음: {input_path}')
        sys.exit(1)

    output_dir.mkdir(parents=True, exist_ok=True)
    print(f'입력: {input_path}')
    print(f'출력: {output_path}\n')

    # ── 전체 시트 로드 ────────────────────────────────────────────────────────
    all_sheets  = pd.read_excel(input_path, sheet_name=None)
    loan_sheets = {k: v for k, v in all_sheets.items() if k != INTEREST_SHEET}

    if INTEREST_SHEET not in all_sheets:
        print(f'[오류] "{INTEREST_SHEET}" 시트를 찾을 수 없습니다.')
        sys.exit(1)

    # ── 차입금별 처리 ─────────────────────────────────────────────────────────
    summary_rows = []

    for sheet_name, raw_df in loan_sheets.items():
        rate = INTEREST_RATES.get(sheet_name)
        if rate is None:
            print(f'[경고] "{sheet_name}" — INTEREST_RATES에 등록되지 않음. 건너뜁니다.\n')
            continue

        df = load_loan_sheet(raw_df)

        # 기초 / 기중 분류
        opening_balance = df.loc[df['_is_bfwd'], '대변'].sum()
        mid_inflow      = df.loc[~df['_is_bfwd'], '대변'].sum()   # 기중 차입
        mid_outflow     = df.loc[~df['_is_bfwd'], '차변'].sum()   # 기중 상환

        balance_daily     = calc_daily_balance(df)
        closing_balance   = balance_daily.iloc[-1]
        expected_interest = calc_expected_interest(balance_daily, rate)

        print(f'■ {sheet_name}')
        print(f'  기초잔액   : {opening_balance:>16,.0f} 원')
        print(f'  기중차입   : {mid_inflow:>16,.0f} 원')
        print(f'  기중상환   : {mid_outflow:>16,.0f} 원')
        print(f'  기말잔액   : {closing_balance:>16,.0f} 원')
        print(f'  적용이자율 : {rate:.4%}')
        print(f'  기대이자   : {expected_interest:>16,.0f} 원\n')

        summary_rows.append({
            '차입금명':     sheet_name,
            '기초잔액':     opening_balance,
            '기중차입':     mid_inflow,
            '기중상환':     mid_outflow,
            '기말잔액':     closing_balance,
            '적용이자율':   rate,
            '기대이자비용': round(expected_interest),
        })

    if not summary_rows:
        print('[종료] 처리된 차입금이 없습니다.')
        sys.exit(0)

    # ── 실제 이자비용 ─────────────────────────────────────────────────────────
    df_int = all_sheets[INTEREST_SHEET].copy()
    df_int['차변'] = pd.to_numeric(df_int['차변'], errors='coerce').fillna(0)
    actual_interest = df_int['차변'].sum()

    # ── 결과 종합 ─────────────────────────────────────────────────────────────
    summary_df     = pd.DataFrame(summary_rows)
    total_expected = summary_df['기대이자비용'].sum()
    difference     = actual_interest - total_expected
    diff_pct       = (difference / actual_interest * 100) if actual_interest else 0.0

    print('=' * 58)
    print(f'  총 기대이자비용  : {total_expected:>16,.0f} 원')
    print(f'  장부상 이자비용  : {actual_interest:>16,.0f} 원')
    print(f'  차       이      : {difference:>16,.0f} 원  ({diff_pct:+.2f}%)')
    print('=' * 58)

    comparison_df = pd.DataFrame([
        {'항목': '총 기대이자비용 (Expected)', '금액(원)': round(total_expected)},
        {'항목': '장부상 이자비용 (Actual)',   '금액(원)': round(actual_interest)},
        {'항목': '차이 (Difference)',          '금액(원)': round(difference)},
        {'항목': '차이율 (%)',                 '금액(원)': round(diff_pct, 2)},
    ])

    # ── 엑셀 저장 ─────────────────────────────────────────────────────────────
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        comparison_df.to_excel(writer, sheet_name='비교', index=False)

        # Summary 열 너비 자동 조정
        ws = writer.sheets['Summary']
        for col in ws.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 50)

    print(f'\n[저장] {output_path}')


if __name__ == '__main__':
    main()
