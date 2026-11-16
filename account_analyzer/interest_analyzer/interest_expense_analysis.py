"""
이자비용 적정성 분석 (Reasonableness Test)
실행: python interest_expense_analysis.py <회사명>
예시: python interest_expense_analysis.py dae_il
입력: interest_analyzer/<회사>/input/이자비용적정성.xlsx  — 차입금별 시트 + '이자비용' 시트
      interest_analyzer/<회사>/input/이자율설정.xlsx    — 회사별 이자율·기간 (아래)
출력: interest_analyzer/<회사>/output/이자비용분석결과.xlsx

이자율설정.xlsx (2026-11-16, 이자율·기간을 코드에서 회사별 파일로 이동 — 예전엔 dae_il 값이
코드에 박혀 있어 다른 회사를 돌리려면 코드를 고쳐야 했음)
  '이자율' 시트: 차입금명 / 연이자율 / 기초잔액 / 비고
    - 차입금명: 이자비용적정성.xlsx의 차입금 시트명과 정확히 같아야 함
    - 연이자율: 0.0362 또는 3.62 둘 다 가능(1보다 크면 %로 보고 100으로 나눔). 빈칸이면 그 차입금 건너뜀
    - 기초잔액(선택): 차입금 시트에 '전기이월' 행이 없을 때(계정 단위 원장이라 차입금별 기초가
      없는 회사, 예: jymts) 입력. 시트의 전기이월 금액에 더해짐
  '설정' 시트: 항목 / 값
    - 기초일, 기말일 (필수) — 일별 잔액 계산 기간
    - 원장기초잔액합계 (선택) — 있으면 차입금별 기초잔액 합계와 대사해 결과 '비교' 시트에 표시
"""

import sys
from pathlib import Path

import numpy as np
import pandas as pd

INTEREST_SHEET = '이자비용'   # 실제 이자비용 시트명
SETTINGS_FILE  = '이자율설정.xlsx'


# ─────────────────────────────────────────────────────────────────────────────

def load_settings(path: Path) -> tuple:
    """이자율설정.xlsx → (rates{차입금명: 연이자율}, openings{차입금명: 기초잔액}, start, end, ledger_open)."""
    xl = pd.ExcelFile(path)
    for s in ('이자율', '설정'):
        if s not in xl.sheet_names:
            raise ValueError(f'{path.name}에 "{s}" 시트가 없습니다.')

    cfg = pd.read_excel(path, sheet_name='설정', header=0)
    cfg = {str(k).strip(): v for k, v in zip(cfg.iloc[:, 0], cfg.iloc[:, 1]) if pd.notna(k)}
    if pd.isna(cfg.get('기초일')) or pd.isna(cfg.get('기말일')):
        raise ValueError(f'{path.name} "설정" 시트에 기초일·기말일이 필요합니다.')
    start, end = pd.Timestamp(cfg['기초일']), pd.Timestamp(cfg['기말일'])
    ledger_open = pd.to_numeric(cfg.get('원장기초잔액합계'), errors='coerce')
    ledger_open = None if pd.isna(ledger_open) else float(ledger_open)

    rt = pd.read_excel(path, sheet_name='이자율')
    rates, openings = {}, {}
    for _, r in rt.iterrows():
        name = str(r.get('차입금명', '')).strip()
        if not name or name == 'nan':
            continue
        rate = pd.to_numeric(r.get('연이자율'), errors='coerce')
        if pd.notna(rate):
            rates[name] = float(rate) / 100 if rate > 1 else float(rate)
        op = pd.to_numeric(r.get('기초잔액'), errors='coerce')
        if pd.notna(op):
            openings[name] = float(op)
    return rates, openings, start, end, ledger_open


def load_loan_sheet(df: pd.DataFrame, period_start: pd.Timestamp) -> pd.DataFrame:
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
    df.loc[df['_is_bfwd'], '날짜'] = period_start
    df = df.dropna(subset=['날짜']).sort_values('날짜').reset_index(drop=True)
    return df


def calc_daily_balance(df: pd.DataFrame, period_start, period_end) -> pd.Series:
    """거래 데이터를 기간 전체 일별 잔액 Series로 변환."""
    df['잔액'] = (df['대변'] - df['차변']).cumsum()
    # 같은 날짜에 거래가 여러 건이면 마지막(당일 종가) 잔액 사용
    balance_by_day = df.groupby('날짜')['잔액'].last()
    date_range = pd.date_range(period_start, period_end, freq='D')
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
    input_path    = company_dir / 'input'  / '이자비용적정성.xlsx'
    settings_path = company_dir / 'input'  / SETTINGS_FILE
    output_dir    = company_dir / 'output'
    output_path   = output_dir  / '이자비용분석결과.xlsx'

    for p in (input_path, settings_path):
        if not p.exists():
            print(f'[오류] 파일 없음: {p}')
            sys.exit(1)

    rates, openings, period_start, period_end, ledger_open = load_settings(settings_path)

    output_dir.mkdir(parents=True, exist_ok=True)
    print(f'입력: {input_path}')
    print(f'설정: {settings_path}  (기간 {period_start:%Y-%m-%d} ~ {period_end:%Y-%m-%d})')
    print(f'출력: {output_path}\n')

    # ── 전체 시트 로드 ────────────────────────────────────────────────────────
    all_sheets  = pd.read_excel(input_path, sheet_name=None)
    loan_sheets = {k: v for k, v in all_sheets.items() if k != INTEREST_SHEET}

    if INTEREST_SHEET not in all_sheets:
        print(f'[오류] "{INTEREST_SHEET}" 시트를 찾을 수 없습니다.')
        sys.exit(1)

    # ── 차입금별 처리 ─────────────────────────────────────────────────────────
    summary_rows = []
    opening_all = 0.0   # 이자율 유무와 관계없이 전체 차입금 기초잔액 합계 (원장 대사용)

    for sheet_name, raw_df in loan_sheets.items():
        if sheet_name in openings:   # 설정 파일의 기초잔액 → 전기이월 행으로 추가
            bfwd = pd.DataFrame([{'날짜': pd.NaT, '적요란': '전기이월', '차변': 0, '대변': openings[sheet_name]}])
            raw_df = bfwd if raw_df.empty else pd.concat([bfwd, raw_df], ignore_index=True)
        df = load_loan_sheet(raw_df, period_start)
        opening_balance = df.loc[df['_is_bfwd'], '대변'].sum() - df.loc[df['_is_bfwd'], '차변'].sum()
        opening_all += opening_balance

        rate = rates.get(sheet_name)
        if rate is None:
            print(f'[경고] "{sheet_name}" — 이자율설정에 연이자율이 없음. 건너뜁니다.\n')
            continue

        # 기초 / 기중 분류
        mid_inflow      = df.loc[~df['_is_bfwd'], '대변'].sum()   # 기중 차입
        mid_outflow     = df.loc[~df['_is_bfwd'], '차변'].sum()   # 기중 상환

        balance_daily     = calc_daily_balance(df, period_start, period_end)
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

    # ── 실제 이자비용 (차변 - 대변: 이자비용 환입·취소분 반영) ─────────────────
    df_int = all_sheets[INTEREST_SHEET].copy()
    df_int['차변'] = pd.to_numeric(df_int['차변'], errors='coerce').fillna(0)
    df_int['대변'] = pd.to_numeric(df_int.get('대변', 0), errors='coerce').fillna(0)
    actual_interest = df_int['차변'].sum() - df_int['대변'].sum()

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

    comparison_rows = [
        {'항목': '총 기대이자비용 (Expected)', '금액(원)': round(total_expected)},
        {'항목': '장부상 이자비용 (Actual)',   '금액(원)': round(actual_interest)},
        {'항목': '차이 (Difference)',          '금액(원)': round(difference)},
        {'항목': '차이율 (%)',                 '금액(원)': round(diff_pct, 2)},
    ]
    # 검증: 차입금별 기초잔액 합계 = 원장 차입금 계정 기초잔액 합계 (설정에 있을 때만)
    if ledger_open is not None:
        gap = opening_all - ledger_open
        ok = '일치' if abs(gap) < 1 else '불일치 — 이자율설정 기초잔액 확인'
        print(f'  [검증] 차입금별 기초잔액 합계 {opening_all:,.0f} vs 원장 {ledger_open:,.0f} → {ok}')
        comparison_rows += [
            {'항목': '[검증] 차입금별 기초잔액 합계', '금액(원)': round(opening_all)},
            {'항목': '[검증] 원장 차입금 기초잔액 합계', '금액(원)': round(ledger_open)},
            {'항목': f'[검증] 차이 ({ok})', '금액(원)': round(gap)},
        ]
    comparison_df = pd.DataFrame(comparison_rows)

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
