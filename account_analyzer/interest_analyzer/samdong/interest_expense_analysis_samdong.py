"""
이자비용 적정성 분석 — samdong 전용
실행: python interest_expense_analysis_samdong.py
입력: journal_analyzer/samdong/data/current/당기_*_분개장_*.xlsx (분개장)
      journal_analyzer/samdong/data/previous/전기_계정별_거래처별_명세_*.xlsx (기초잔액)
출력: account_analyzer/interest_analyzer/samdong/output/이자비용분석결과_samdong.xlsx

dae_il용 interest_expense_analysis.py 와 다른 점
-------------------------------------------------
dae_il은 감사인이 손으로 정리한 거래처원장(날짜/적요란/차변/대변)을 input 폴더에
직접 넣고, 이자율도 INTEREST_RATES 딕셔너리에 감사인이 직접 입력하는 구조다.

samdong은 ERP 원장의 "관리항목" 슬롯(계정과목마다 의미가 다른 가변 컬럼)에
차입금 관리번호와 "그 시점 실제 적용금리"가 이미 통째로 찍혀 있고, 변동금리라
월마다 금리가 바뀐다(예: #10693 대출이 5.226% → 4.816% → 4.516% → 4.316%로
연중 변동). 그래서:
  - 감사인이 이자율을 입력할 필요가 없다 — 원장에 찍힌 실제 적용금리를 그대로 쓴다.
  - 단일 연이자율이 아니라, 이자비용 지급 구간(전표 간격)마다 그때 적용된 금리로
    "그 구간 평균잔액 × 금리 × 구간일수/365"를 계산해 합산한다.

관리항목 슬롯 매핑(2026-09-08 실제 데이터로 확인, samdong 원장 전용)
--------------------------------------------------------------------
차입금 라인 (단기차입금-구매자금/일반대출, 장기차입금, 유동성장기부채-원화):
  관리항목1 = "#관리번호 대출종류" (예: "#10925 기업구매자금")
  관리항목2 = 은행/지점
이자비용 라인 (이자비용-구매자금/일반대출; 대출보증료인 이자비용-대출보증은 관리번호가
없는 별개 항목이라 제외):
  관리항목1 = 계좌번호
  관리항목2 = 은행/지점
  관리항목3 = "#관리번호 대출종류"
  관리항목4 = 그 지급 시점 적용금리(%)
사채/유동성사채는 관리번호 형식이 아니라(채권명) 애초에 대상 아님.

관리번호는 계정과목(단기/장기/유동성장기부채)이 바뀌어도(유동성대체) 그대로
유지되므로, 잔액 누적 시 계정과목을 가리지 않고 관리번호로만 필터링한다.

테스트 대상에서 제외: 구매자금대출, 만기 차환(재대출) 건
------------------------------------------------------
- 구매자금대출(단기차입금-구매자금/이자비용-구매자금)은 매입 건별로 연중 수시 실행·
  상환되며 그때마다 관리번호가 새로 채번돼(#10925 → #10926 → #10928 ...) 관리번호
  하나=대출 하나로 추적이 안 됨(사용자 확인, 2026-09-08). LOAN_ACCOUNTS/
  INTEREST_ACCOUNTS에서 아예 제외.
- 일반대출류도 만기에 차환(재대출)되면서 동일 텍스트 관리번호가 실제로는 다른
  은행 계좌(계좌번호가 바뀜)를 가리키는 경우가 있어(예: #92010101이 전기말엔 계좌
  "722406792-01-01-01", 당기 차환 후엔 "722506792-01-01-01") 기초잔액이 새
  관리번호로 안 넘어오거나 이중 계산될 수 있다. 이런 건은 자동 보정하지 않고
  Summary의 '확인필요' 플래그(기말잔액 음수, 기대이자합계 음수, |차이율|>
  FLAG_PCT_THRESHOLD)로 걸러 '확인필요' 시트에 모아 감사인이 별도 확인하게 한다
  — 회사 전체 비교(비교 시트)도 이 건들은 제외하고 집계한다.
"""

import os
import re
from pathlib import Path

import pandas as pd

REPO_ROOT   = Path(__file__).resolve().parents[3]
SAMDONG_DIR = REPO_ROOT / 'journal_analyzer' / 'samdong' / 'data'
OUTPUT_PATH = Path(__file__).parent / 'output' / '이자비용분석결과_samdong.xlsx'

# 구매자금대출(단기차입금-구매자금/이자비용-구매자금)은 실물 매입 건별로 연중 수시로
# 실행·상환되며 그때마다 관리번호가 새로 채번되어(#10925, #10926, #10928... 등 계속
# 증가) 관리번호 하나=대출 하나로 추적이 안 된다(사용자 확인, 2026-09-08). 이 테스트는
# 관리번호가 연중 안정적으로 유지되는 일반대출/시설자금류만 대상으로 한다.
LOAN_ACCOUNTS     = ['단기차입금-일반대출', '장기차입금', '유동성장기부채-원화']
INTEREST_ACCOUNTS = ['이자비용-일반대출']
FLAG_PCT_THRESHOLD = 30  # |차이율(%)| 이 이 값을 넘으면 확인필요로 플래그

_ID_RE   = re.compile(r'^(\S+)')
_DATE_RE = re.compile(r'(20\d{6})')


def _find_file(dir_path, must_contain):
    if not os.path.isdir(dir_path):
        return None
    for f in sorted(os.listdir(dir_path)):
        if f.startswith('~$') or not f.endswith('.xlsx'):
            continue
        if all(k in f for k in must_contain):
            return os.path.join(dir_path, f)
    return None


def _split_id(text):
    """'#10925 기업구매자금' → ('#10925', '기업구매자금')"""
    s = str(text).strip()
    m = _ID_RE.match(s)
    if not m:
        return None, ''
    mgmt_id = m.group(1)
    rest = s[len(mgmt_id):].strip()
    return mgmt_id, rest


def load_journal():
    path = _find_file(SAMDONG_DIR / 'current', ['당기', '분개장'])
    if not path:
        raise FileNotFoundError(f'당기 분개장 파일을 찾을 수 없음: {SAMDONG_DIR / "current"}')
    print(f'[분개장] {path}')
    df = pd.read_excel(path)
    df.columns = [str(c).strip() for c in df.columns]

    df['전표일자'] = pd.to_datetime(
        df['전표승인번호'].astype(str).str.extract(_DATE_RE, expand=False),
        format='%Y%m%d', errors='coerce')
    df['차변금액'] = pd.to_numeric(df['차변금액'], errors='coerce').fillna(0)
    df['대변금액'] = pd.to_numeric(df['대변금액'], errors='coerce').fillna(0)
    df['계정과목'] = df['계정과목'].astype(str).str.strip()
    return df


def load_opening_balances():
    """전기 계정별_거래처별_명세: 단기차입금/장기차입금/유동성장기부채 시트에서
    관리번호별 전기 기말잔액(=당기 기초잔액), 은행명, 종류를 뽑는다."""
    path = _find_file(SAMDONG_DIR / 'previous', ['전기', '계정별', '거래처별', '명세'])
    if not path:
        raise FileNotFoundError(f'전기 계정별_거래처별_명세 파일을 찾을 수 없음: {SAMDONG_DIR / "previous"}')
    print(f'[전기명세] {path}')
    xl = pd.ExcelFile(path)

    opening, labels = {}, {}
    for sheet in ['단기차입금', '장기차입금', '유동성장기부채']:
        if sheet not in xl.sheet_names:
            continue
        sdf = xl.parse(sheet, header=0)
        sdf.columns = [str(c).strip() for c in sdf.columns]
        bal_col = next((c for c in sdf.columns if '잔액' in c or c == '기말'), None)
        if bal_col is None or '관리번호' not in sdf.columns:
            continue
        for _, r in sdf.iterrows():
            mgmt_id = str(r.get('관리번호', '')).strip()
            if not mgmt_id or mgmt_id == 'nan':
                continue
            if '구매자금' in str(r.get('종류', '')):  # 구매자금대출은 테스트 대상 아님
                continue
            bal = pd.to_numeric(r.get(bal_col), errors='coerce')
            opening[mgmt_id] = opening.get(mgmt_id, 0) + (0 if pd.isna(bal) else bal)
            labels.setdefault(mgmt_id, {
                '은행명': str(r.get('은행명', '')).strip(),
                '종류':   str(r.get('종류', '')).strip(),
            })
    return opening, labels


def build_loan_ledger(df):
    """관리번호 → 차입금 라인(날짜, 차변, 대변) DataFrame. 계정과목(유동성대체) 안 가리고 합침."""
    sub = df[df['계정과목'].isin(LOAN_ACCOUNTS)].copy()
    ids, kinds, banks = [], [], []
    for v1, v2 in zip(sub['관리항목1'], sub.get('관리항목2', '')):
        mgmt_id, kind = _split_id(v1)
        ids.append(mgmt_id); kinds.append(kind); banks.append(str(v2).strip())
    sub['관리번호'] = ids
    sub['종류']    = kinds
    sub['은행명']  = banks
    sub = sub.dropna(subset=['관리번호', '전표일자'])
    return sub


def build_interest_lines(df):
    """관리번호 → 이자비용 라인(지급일, 적용금리, 실제이자) DataFrame."""
    sub = df[df['계정과목'].isin(INTEREST_ACCOUNTS)].copy()
    ids, kinds = [], []
    for v3 in sub.get('관리항목3', ''):
        mgmt_id, kind = _split_id(v3)
        ids.append(mgmt_id); kinds.append(kind)
    sub['관리번호'] = ids
    sub['종류']    = kinds
    sub['적용금리'] = pd.to_numeric(sub.get('관리항목4'), errors='coerce')
    sub = sub.dropna(subset=['관리번호', '전표일자'])
    return sub.sort_values('전표일자')


def calc_daily_balance(loan_sub, opening_bal, period_start, period_end):
    d = loan_sub.sort_values('전표일자').copy()
    d['잔액증감'] = d['대변금액'] - d['차변금액']
    by_day = d.groupby('전표일자')['잔액증감'].sum().sort_index()
    date_range = pd.date_range(period_start, period_end, freq='D')
    daily_change = by_day.reindex(date_range, fill_value=0)
    return opening_bal + daily_change.cumsum()


def analyze():
    df = load_journal()
    opening, labels = load_opening_balances()

    valid_year = df['전표일자'].dt.year.dropna()
    year = int(valid_year.mode().iloc[0])
    period_start, period_end = pd.Timestamp(year, 1, 1), pd.Timestamp(year, 12, 31)
    print(f'[기간] {period_start.date()} ~ {period_end.date()}')

    loan_all = build_loan_ledger(df)
    int_all  = build_interest_lines(df)

    all_ids = sorted(set(opening) | set(loan_all['관리번호']) | set(int_all['관리번호']))
    print(f'[관리번호] 총 {len(all_ids)}건')

    detail_rows, summary_rows = [], []

    for mgmt_id in all_ids:
        loan_sub = loan_all[loan_all['관리번호'] == mgmt_id]
        int_sub  = int_all[int_all['관리번호'] == mgmt_id]

        label = labels.get(mgmt_id, {})
        bank = label.get('은행명', '')
        kind = label.get('종류', '')
        if not bank and not loan_sub.empty:
            bank = loan_sub['은행명'].iloc[-1]
        if not kind:
            if not loan_sub.empty:
                kind = loan_sub['종류'].iloc[-1]
            elif not int_sub.empty:
                kind = int_sub['종류'].iloc[-1]

        opening_bal = opening.get(mgmt_id, 0)
        daily_balance = calc_daily_balance(loan_sub, opening_bal, period_start, period_end)
        closing_bal = daily_balance.iloc[-1]

        prev_date = period_start
        total_expected, total_actual, no_rate_cnt = 0.0, 0.0, 0
        for _, r in int_sub.iterrows():
            pay_date = r['전표일자']
            if pd.isna(pay_date) or pay_date < prev_date:
                continue
            days = (pay_date - prev_date).days + 1
            avg_bal = daily_balance.loc[prev_date:pay_date].mean()
            rate = r['적용금리']
            actual = r['차변금액']
            if pd.isna(rate):
                expected = 0.0
                no_rate_cnt += 1
            else:
                expected = float(avg_bal) * (rate / 100) * days / 365

            detail_rows.append({
                '관리번호': mgmt_id, '은행명': bank, '종류': kind,
                '계정과목': r['계정과목'], '지급일자': pay_date.date(),
                '기간시작': prev_date.date(), '기간일수': days,
                '기간평균잔액': round(float(avg_bal)),
                '적용금리(%)': rate if pd.notna(rate) else '',
                '기대이자': round(expected), '실제이자': round(float(actual)),
                '차이': round(float(actual) - expected),
            })
            total_expected += expected
            total_actual   += actual
            prev_date = pay_date + pd.Timedelta(days=1)

        diff = total_actual - total_expected
        diff_pct = (diff / total_actual * 100) if total_actual else 0.0

        # 차환(재대출)으로 관리번호가 재사용/재채번되면 기초잔액이 새 관리번호로
        # 안 넘어오거나 이중계산되어 기대이자·기말잔액이 말이 안 되는 값으로 튄다.
        # 이런 케이스는 자동 보정하지 않고 감사인 확인용으로 플래그만 남긴다.
        issues = []
        if no_rate_cnt:
            issues.append(f'금리정보없음 {no_rate_cnt}건')
        suspect_rollover = closing_bal < 0 or (total_actual > 0 and total_expected < 0)
        if suspect_rollover:
            issues.append('관리번호 재사용/차환 의심(잔액 계산 이상)')
        large_diff = bool(total_actual) and abs(diff_pct) > FLAG_PCT_THRESHOLD
        if large_diff and not suspect_rollover:
            issues.append(f'차이율 {diff_pct:.1f}% 큼')
        need_review = suspect_rollover or large_diff

        summary_rows.append({
            '관리번호': mgmt_id, '은행명': bank, '종류': kind,
            '기초잔액': round(opening_bal), '기말잔액': round(float(closing_bal)),
            '이자지급건수': len(int_sub), '기대이자합계': round(total_expected),
            '실제이자합계': round(total_actual), '차이': round(diff),
            '차이율(%)': round(diff_pct, 2),
            '확인필요': 'Y' if need_review else '',
            '비고': '; '.join(issues),
        })

    summary_df = pd.DataFrame(summary_rows)
    summary_df = (summary_df.assign(_abs=summary_df['차이'].abs())
                  .sort_values('_abs', ascending=False).drop(columns=['_abs'])
                  .reset_index(drop=True))
    detail_df = pd.DataFrame(detail_rows)
    review_df = summary_df[summary_df['확인필요'] == 'Y'].reset_index(drop=True)

    # 회사 전체 비교는 확인필요(관리번호 재사용/차환 의심 등)로 플래그된 건을 제외하고
    # 집계 — 그 건들의 계산값 자체를 신뢰할 수 없어 섞으면 전체 비교가 왜곡된다.
    ok_df = summary_df[summary_df['확인필요'] != 'Y']
    total_expected = ok_df['기대이자합계'].sum()
    total_actual   = ok_df['실제이자합계'].sum()
    total_diff     = total_actual - total_expected
    total_diff_pct = (total_diff / total_actual * 100) if total_actual else 0.0
    comparison_df = pd.DataFrame([
        {'항목': '총 기대이자비용 (Expected, 확인필요 건 제외)', '금액(원)': round(total_expected)},
        {'항목': '장부상 이자비용 (Actual, 확인필요 건 제외)',   '금액(원)': round(total_actual)},
        {'항목': '차이 (Difference)',                            '금액(원)': round(total_diff)},
        {'항목': '차이율 (%)',                                   '금액(원)': round(total_diff_pct, 2)},
        {'항목': '확인필요로 제외된 관리번호 수',                 '금액(원)': len(review_df)},
    ])

    print('=' * 58)
    print(f'  총 기대이자비용  : {total_expected:>16,.0f} 원')
    print(f'  장부상 이자비용  : {total_actual:>16,.0f} 원')
    print(f'  차       이      : {total_diff:>16,.0f} 원  ({total_diff_pct:+.2f}%)')
    print('=' * 58)

    print(f'  확인필요(제외)   : {len(review_df):>16} 건')
    print('=' * 58)

    OUTPUT_PATH.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(OUTPUT_PATH, engine='openpyxl') as writer:
        summary_df.to_excel(writer, sheet_name='Summary', index=False)
        review_df.to_excel(writer, sheet_name='확인필요', index=False)
        detail_df.to_excel(writer, sheet_name='이자비용_상세', index=False)
        comparison_df.to_excel(writer, sheet_name='비교', index=False)
        for sname in ['Summary', '확인필요', '이자비용_상세', '비교']:
            ws = writer.sheets[sname]
            for col in ws.columns:
                max_len = max(len(str(c.value or '')) for c in col)
                ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 50)

    print(f'\n[저장] {OUTPUT_PATH}')
    return summary_df, detail_df, comparison_df


if __name__ == '__main__':
    analyze()
