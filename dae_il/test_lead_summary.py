# -*- coding: utf-8 -*-
import pandas as pd
import os
import sys

script_dir = os.path.dirname(os.path.abspath(__file__))

coa_path = os.path.join(script_dir, 'coa_dae_il_25.xlsx')
bal_path = os.path.join(script_dir, 'results', 'dae_il_20260504_일반사항분석.xlsx')
out_path = os.path.join(script_dir, 'test_lead_summary.xlsx')
log_path = os.path.join(script_dir, 'test_lead_summary_log.txt')

log_lines = []
def log(msg=''):
    log_lines.append(msg)
    print(msg)

# ── 1. 데이터 로드 ──────────────────────────────────────────
df_coa = pd.read_excel(coa_path)
df_bal = pd.read_excel(bal_path)

log(f'COA 행 수: {len(df_coa)},  잔액 행 수: {len(df_bal)}')

# ── 2. 계정코드 문자열 통일 ────────────────────────────────
df_coa['계정코드'] = df_coa['계정코드'].astype(str)
df_bal['계정코드'] = df_bal['계정코드'].astype(str)

# ── 3. Left Join 병합 ─────────────────────────────────────
merge_cols = ['계정코드', '대분류 (Level 1)', '리드스케줄 (Level 4)', '계정성격']
df = pd.merge(df_bal, df_coa[merge_cols], on='계정코드', how='left')

missing = df['대분류 (Level 1)'].isna().sum()
if missing > 0:
    log(f'[경고] COA에 매핑 안 된 계정: {missing}개')
    log(df[df['대분류 (Level 1)'].isna()][['계정코드', '계정과목']].to_string())

# ── 4. 차감 계정 순액 처리 ─────────────────────────────────
df['순기말잔액'] = df['기말잔액'].copy()
df['순기초잔액'] = df['기초잔액'].copy()
df['순차변합계'] = df['차변 합계'].copy()
df['순대변합계'] = df['대변 합계'].copy()

mask_contra = df['계정성격'].str.contains('차감', na=False)
df.loc[mask_contra, '순기말잔액'] = df.loc[mask_contra, '기말잔액'] * -1
df.loc[mask_contra, '순기초잔액'] = df.loc[mask_contra, '기초잔액'] * -1
df.loc[mask_contra, '순차변합계'] = df.loc[mask_contra, '차변 합계'] * -1
df.loc[mask_contra, '순대변합계'] = df.loc[mask_contra, '대변 합계'] * -1

log(f'\n차감 계정 적용: {mask_contra.sum()}개')
log(df[mask_contra][['계정코드','계정과목','계정성격']].to_string())

# ── 5. 그룹화 ─────────────────────────────────────────────
grp = df.groupby(
    ['대분류 (Level 1)', '리드스케줄 (Level 4)'],
    dropna=False
).agg(
    기초잔액=('순기초잔액', 'sum'),
    차변합계=('순차변합계', 'sum'),
    대변합계=('순대변합계', 'sum'),
    기말잔액_순액=('순기말잔액', 'sum')
).reset_index()

grp.columns = ['대분류', '리드스케줄', '기초잔액', '차변합계', '대변합계', '기말잔액(순액)']

# ── 6. 결과 출력 ──────────────────────────────────────────
pd.set_option('display.float_format', '{:,.0f}'.format)
pd.set_option('display.max_rows', 100)
pd.set_option('display.width', 200)

log('\n======== 리드스케줄별 잔액 합계 ========')
log(grp.to_string(index=False))

# ── 7. 엑셀 저장 ──────────────────────────────────────────
with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
    grp.to_excel(writer, sheet_name='리드스케줄별합계', index=False)

    detail_cols = ['계정코드', '계정과목', '대분류 (Level 1)', '리드스케줄 (Level 4)',
                   '계정성격', '기초잔액', '차변 합계', '대변 합계', '기말잔액', '순기말잔액']
    df[detail_cols].to_excel(writer, sheet_name='계정별상세', index=False)

log(f'\n저장 완료: {out_path}')
log('  - 시트1: 리드스케줄별합계')
log('  - 시트2: 계정별상세')

# ── 8. 로그 파일 저장 ─────────────────────────────────────
with open(log_path, 'w', encoding='utf-8') as f:
    f.write('\n'.join(log_lines))
