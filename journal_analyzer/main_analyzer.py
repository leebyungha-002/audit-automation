#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
journal_analyzer/main_analyzer.py  v3
analysis.py 20개 분석 메뉴 — DataFrame 반환 방식

폴더 구조:
  journal_analyzer/
  ├── main_analyzer.py
  └── {company}/
      ├── task_list_{company}.xlsx      ← 실행 제어 + 파라미터
      ├── data/
      │   ├── current/                  ← 당기 분개장 (csv/xlsx)
      │   └── previous/                 ← 전기 분개장 (csv/xlsx)
      └── results/                      ← 분석 결과 저장

task_list 파라미터 시트 규격:
  거래처비교   : 계정과목 / 금액열(차변·대변·both) / 실행여부
  벤포드분석   : 계정과목 / 금액열 / 실행여부
  일자차이분석 : 기준일수 / 실행여부
  상대계정분석 : 계정과목 / 금액열 / 실행여부
  키워드검색   : 키워드 / 실행여부
  라운드넘버   : 계정과목 / 금액열 / 최소금액 / 실행여부
  특수관계자   : 거래처명 / 실행여부
  자산부채교차 : 구분(자산·부채) / 계정과목 / 실행여부
  매출비용교차 : 구분(매출·비용) / 계정과목 / 실행여부
  심층분석     : 계정과목 / 개수 / 금액열(차변·대변·both) / 실행여부
  AI계정별분석 : 계정과목 / 실행여부
  거래처분석   : 작업명 / 계정과목 / 거래처명 / 금액열 / 실행여부
  벤포드이탈   : 계정과목 / 금액열 / 임계값 / 최대건수 / 실행여부

실행:  python main_analyzer.py sejoong
"""

# =============================================================================
# 0. Imports & 상수
# =============================================================================
import sys, os, glob, re, io, argparse, warnings, unicodedata

# Windows 터미널 한글 깨짐 방지 (cp949 → UTF-8 강제)
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
if hasattr(sys.stderr, 'reconfigure'):
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')
import pandas as pd
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import openpyxl
from openpyxl.drawing.image import Image as XLImage
from openpyxl.utils.dataframe import dataframe_to_rows

warnings.filterwarnings('ignore', category=pd.errors.DtypeWarning)

BASE_DIR          = os.path.dirname(os.path.abspath(__file__))
TASK_MASTER_SHEET = '분석목록'

# 컬럼명 상수
COL_DATE       = '전표일자'
COL_JOURNAL_ID = '전표번호'
COL_ACCOUNT    = '계정명'
COL_CLIENT     = '거래처명'
COL_DESC       = '적요'
COL_DEBIT      = '차변'
COL_CREDIT     = '대변'
COL_EMPLOYEE   = '사원명'

# 분석 상수
BENFORD_MIN_ROWS        = 5
BENFORD_PROBS           = {1:0.301,2:0.176,3:0.125,4:0.097,5:0.079,6:0.067,7:0.058,8:0.051,9:0.046}
DEFAULT_BENFORD_TARGETS = [('복리후생비','차변'),('접대비','차변'),('여비교통비','차변')]
DEFAULT_KEYWORDS        = ['상품권','접대','가수금','선물','회식','결산수정','비자금','가지급','리베이트','현금']
MASK_TARGET_COLS        = ['적요','관리항목4','비고','내용']
GLOBAL_SAFE_MAP: dict   = {}
_CLIENT_COUNTER         = 1
_COMPANY_DIR            = None


# =============================================================================
# 1. 유틸 함수
# =============================================================================

def _nv(val, blank_vals=('(전체)','nan','none','')) -> str:
    s = str(val).strip()
    return '' if s.lower() in blank_vals else s

def _safe_float(val, default: float) -> float:
    try:
        f = float(val)
        return default if pd.isna(f) else f
    except (TypeError, ValueError):
        return default

def get_first_digit(number):
    try:
        s = str(abs(int(number)))
        return int(s[0]) if s[0] != '0' else 0
    except Exception:
        return 0

def _has_expected_columns(df):
    cols = [str(c).strip() for c in df.columns]
    keywords = ('차변','대변','계정','전표일자','거래처','전표번호','적요')
    return any(any(kw in c for kw in keywords) for c in cols)

def _read_excel_with_header_detection(path, dtype_map=None):
    dtype_map = dtype_map or {}
    for header_row in range(3):
        try:
            trial = pd.read_excel(path, engine='openpyxl', header=header_row, dtype=dtype_map)
            if trial.empty or len(trial.columns) < 2:
                continue
            if _has_expected_columns(trial):
                return trial, header_row
        except Exception:
            continue
    df = pd.read_excel(path, engine='openpyxl', dtype=dtype_map)
    return df, 0


def _load_with_parquet_cache(path: str, dtype_map: dict = None) -> tuple:
    """Excel 파일을 읽되, Parquet 캐시가 최신이면 캐시에서 빠르게 로드."""
    dtype_map = dtype_map or {}
    cache_dir = os.path.join(os.path.dirname(os.path.dirname(path)), 'cache')
    cache_name = os.path.splitext(os.path.basename(path))[0] + '.parquet'
    cache_path = os.path.join(cache_dir, cache_name)

    if os.path.isfile(cache_path) and os.path.getmtime(cache_path) >= os.path.getmtime(path):
        print(f'     ⚡ 캐시 로드: {cache_name}', flush=True)
        df = pd.read_parquet(cache_path)
        for col, dt in dtype_map.items():
            if col in df.columns:
                df[col] = df[col].astype(dt)
        return df, 0

    df, header_row = _read_excel_with_header_detection(path, dtype_map)
    try:
        os.makedirs(cache_dir, exist_ok=True)
        df.to_parquet(cache_path, index=False)
        print(f'     💾 캐시 저장: {cache_name}', flush=True)
    except Exception as e:
        print(f'     ⚠️ 캐시 저장 실패: {e}', flush=True)
    return df, header_row

def _normalize_account_for_match(s):
    s = str(s)
    for full, half in [('（','('),('）',')'),('⦅','('),('⦆',')'),
                       ('﹙','('),('﹚',')'),('【','('),('】',')')]:
        s = s.replace(full, half)
    s = s.replace(' ',' ').replace('＆','&')
    s = re.sub(r'\s+','',s)
    s = re.sub(r'[()（）]+','',s)
    s = s.replace('권','').lower()
    return s

def _account_match_flexible(acct_series, acct_str):
    acct_str = str(acct_str).strip()
    if not acct_str:
        return pd.Series(False, index=acct_series.index)
    norm_series = acct_series.fillna('').astype(str).apply(_normalize_account_for_match)
    norm_user   = _normalize_account_for_match(acct_str)
    mask = norm_series.str.startswith(norm_user)
    if mask.any(): return mask
    mask = norm_series.str.contains(re.escape(norm_user), na=False, regex=True, case=False)
    if mask.any(): return mask
    def _row(val):
        nv = _normalize_account_for_match(str(val) if pd.notna(val) else '')
        if not nv: return False
        return norm_user in nv or nv in norm_user
    return acct_series.fillna('').apply(_row)

def _to_numeric_amount(series):
    s = series.astype(str).str.strip()
    s = s.str.replace(',','',regex=False).str.replace(' ','',regex=False)
    s = s.str.replace(r'^#+$','0',regex=True)
    return pd.to_numeric(s, errors='coerce').fillna(0)

def _normalize_date_journal_columns(df):
    df.columns = [str(c).strip() for c in df.columns]
    if COL_DATE not in df.columns:
        for c in df.columns:
            if c in ('일자','전표일자','전표일','거래일자','적요일자'):
                df = df.rename(columns={c: COL_DATE}); break
    if COL_JOURNAL_ID not in df.columns:
        for c in df.columns:
            if c in ('전표등록번호','전표번호','전표NO','전표 no'):
                df = df.rename(columns={c: COL_JOURNAL_ID}); break
    if COL_CLIENT not in df.columns:
        for c in df.columns:
            if c in ('거래처명','거래처','상대거래처','거래처명칭'):
                df = df.rename(columns={c: COL_CLIENT}); break
    if COL_ACCOUNT not in df.columns:
        for c in df.columns:
            if c in ('계정명','계정과목','계정','과목'):
                df = df.rename(columns={c: COL_ACCOUNT}); break
    return df

def _normalize_debit_credit_columns(df):
    df.columns = [str(c).strip() for c in df.columns]
    debit_cols = [c for c in df.columns if '차변' in c]
    if debit_cols:
        for c in debit_cols: df[c] = _to_numeric_amount(df[c])
        primary = COL_DEBIT if COL_DEBIT in debit_cols else debit_cols[0]
        df[COL_DEBIT] = df[primary].copy()
        for c in debit_cols:
            if c != COL_DEBIT:
                df[COL_DEBIT] = df[COL_DEBIT].fillna(df[c])
                df = df.drop(columns=[c], errors='ignore')
    credit_cols = [c for c in df.columns if '대변' in c and c != COL_DEBIT]
    if credit_cols:
        for c in credit_cols: df[c] = _to_numeric_amount(df[c])
        primary = COL_CREDIT if COL_CREDIT in credit_cols else credit_cols[0]
        df[COL_CREDIT] = df[primary].copy()
        for c in credit_cols:
            if c != COL_CREDIT:
                df[COL_CREDIT] = df[COL_CREDIT].fillna(df[c])
                df = df.drop(columns=[c], errors='ignore')
    return df

def _preprocess_df(df):
    df = df.copy()
    df.columns = df.columns.str.strip()
    df = _normalize_date_journal_columns(df)
    df = _normalize_debit_credit_columns(df)
    if COL_DATE in df.columns:
        ser = df[COL_DATE]
        try:
            num = pd.to_numeric(ser, errors='coerce')
            if num.notna().sum() > len(ser) * 0.5:
                ser = num.fillna(0).astype('int64').astype(str)
                df[COL_DATE] = pd.to_datetime(ser, format='%Y%m%d', errors='coerce')
            else:
                df[COL_DATE] = pd.to_datetime(ser, errors='coerce')
        except Exception:
            df[COL_DATE] = pd.to_datetime(ser, errors='coerce')
    if COL_DEBIT  in df.columns: df[COL_DEBIT]  = _to_numeric_amount(df[COL_DEBIT])
    if COL_CREDIT in df.columns: df[COL_CREDIT] = _to_numeric_amount(df[COL_CREDIT])
    reg_names = ['등록일자','등록일','작성일자','작성일','생성일자','생성일','입력일자','입력일']
    for col in df.columns:
        if any(n in str(col).strip() for n in reg_names):
            try:
                ser_r = df[col]
                num_r = pd.to_numeric(ser_r, errors='coerce')
                if num_r.notna().sum() > len(ser_r) * 0.5:
                    df[col] = pd.to_datetime(num_r.fillna(0).astype('int64').astype(str), format='%Y%m%d', errors='coerce')
                else:
                    df[col] = pd.to_datetime(ser_r, errors='coerce')
            except Exception:
                df[col] = pd.to_datetime(df[col], errors='coerce')
            break
    return df

def _get_gubun_col(df):
    for c in df.columns:
        if str(c).strip() == '구분': return c
    return None

def get_safe_client_name(real_name):
    global _CLIENT_COUNTER
    if pd.isna(real_name) or str(real_name).strip() == '': return '(미기재)'
    real_name = str(real_name).strip()
    if real_name in GLOBAL_SAFE_MAP: return GLOBAL_SAFE_MAP[real_name]
    alias = f'Client_{_CLIENT_COUNTER:03d}'
    GLOBAL_SAFE_MAP[real_name] = alias
    _CLIENT_COUNTER += 1
    return alias

def mask_sensitive_info(text):
    if not isinstance(text, str):
        text = str(text) if text is not None else ''
        if text == '': return text
    text = re.sub(r'(\d{2,})[-](\d{2,})[-](\d{2,})', r'\1-****-\3', text)
    def _mask(m):
        s = m.group()
        return s[:4] + '****' + s[-4:] if len(s) >= 10 else s
    return re.sub(r'\b\d{10,}\b', _mask, text)

def _related_party_pattern(party):
    p = str(party).strip()
    if p.startswith('(주)'): return r'(?:\(주\))?' + re.escape(p[3:])
    return re.escape(p)

def draw_benford_chart(account_name, direction, actual_probs, benford_probs):
    try:
        plt.rc('font', family='Malgun Gothic')
        plt.rcParams['axes.unicode_minus'] = False
        digits = range(1, 10)
        plt.figure(figsize=(10, 6))
        plt.plot(digits, [benford_probs[d]*100 for d in digits],
                 color='red', marker='o', linestyle='--', label='벤포드 법칙')
        plt.bar(digits, [actual_probs.get(d,0.0)*100 for d in digits],
                color='skyblue', alpha=0.7, label=f'실제 ({account_name})')
        plt.title(f'벤포드 분석: {account_name} ({direction})')
        plt.legend(); plt.grid(axis='y', linestyle='--', alpha=0.5)
        buf = io.BytesIO()
        plt.savefig(buf, format='png'); plt.close(); buf.seek(0)
        return buf
    except Exception:
        return None

def _safe_sheet(name, max_len=31):
    s = re.sub(r'[\\/*?:\[\]]', '', str(name).strip())
    return s[:max_len]


# =============================================================================
# 2. 회사별 전용 전처리 모듈 동적 로드
# =============================================================================

def _apply_company_preprocess(df: pd.DataFrame, company_name: str) -> pd.DataFrame:
    """
    journal_analyzer/{company}/preprocess.py 가 존재하면 동적으로 임포트하여
    preprocess(df) 함수를 실행한다. 파일이 없으면 df 를 그대로 반환한다.

    규약
    ----
    각 회사 preprocess.py 는 반드시 아래 시그니처를 가져야 한다.
      def preprocess(df: pd.DataFrame) -> pd.DataFrame
    """
    import importlib.util

    module_path = os.path.join(BASE_DIR, company_name, 'preprocess.py')
    if not os.path.isfile(module_path):
        return df   # 전처리 모듈 없음 → 범용 로직으로 진행

    spec   = importlib.util.spec_from_file_location(f'{company_name}.preprocess', module_path)
    module = importlib.util.module_from_spec(spec)
    try:
        spec.loader.exec_module(module)
    except Exception as e:
        print(f'  [전처리 모듈 로드 오류] {company_name}/preprocess.py — {e}')
        return df

    fn = getattr(module, 'preprocess', None)
    if not callable(fn):
        print(f'  [경고] {company_name}/preprocess.py 에 preprocess() 함수가 없습니다.')
        return df

    print(f'  [전처리] {company_name}/preprocess.py 적용 중...')
    return fn(df)


# =============================================================================
# 3. 데이터 로드 (data/current → 당기, data/previous → 전기)
# =============================================================================

def load_data(company_dir: str) -> pd.DataFrame:
    dtype_map   = {COL_JOURNAL_ID: str}
    current_dir  = os.path.join(company_dir, 'data', 'current')
    previous_dir = os.path.join(company_dir, 'data', 'previous')
    all_dfs = []

    def _read(path, fname):
        ext = os.path.splitext(path)[1].lower()
        try:
            if ext == '.csv':
                print(f'     📂 CSV: {fname}', flush=True)
                for enc in ['utf-8','cp949','utf-16','euc-kr','latin-1']:
                    try:
                        out = pd.read_csv(path, encoding=enc, dtype=dtype_map)
                        print(f'     ✅ {len(out)}행', flush=True)
                        return out
                    except UnicodeDecodeError:
                        continue
            elif ext == '.xlsx':
                print(f'     📂 엑셀: {fname}', flush=True)
                out, h = _load_with_parquet_cache(path, dtype_map)
                if h > 0: print(f'     ℹ️ 헤더: {h+1}행', flush=True)
                print(f'     ✅ {len(out)}행', flush=True)
                return out
        except Exception as e:
            print(f'     ⚠️ 로드 실패: {fname} — {e}', flush=True)
        return None

    for label, dir_path in [('당기', current_dir), ('전기', previous_dir)]:
        if not os.path.isdir(dir_path):
            print(f'   ℹ️ 폴더 없음 ({label}): {dir_path}')
            continue
        for f in sorted(os.listdir(dir_path)):
            if not (f.endswith('.csv') or f.endswith('.xlsx')): continue
            if f.startswith('~$'): continue
            df = _read(os.path.join(dir_path, f), f)
            if df is not None and not df.empty:
                df['구분'] = label
                all_dfs.append(df)
                print(f'   📂 [{label}] {f} ({len(df)}행)')

    if not all_dfs: return None
    return pd.concat(all_dfs, axis=0, ignore_index=True)


# =============================================================================
# 3. 분석 함수 (2~19번)
#    시그니처: (df: DataFrame, params_list: list[dict]) → DataFrame | dict[str, DataFrame]
#    dict 반환 시 키 = 시트명
# =============================================================================

# ── 2. 거래처 전기/당기 비교 ──────────────────────────────────────────────────
def analyze_client_comparison(df: pd.DataFrame, params_list: list) -> dict:
    # 계정별 방향 매핑 (금액열 → 구분 순으로 fallback)
    acct_dir = {}
    for p in params_list:
        acct = _nv(p.get('계정과목',''))
        if not acct: continue
        vtype = (_nv(p.get('금액열',''), blank_vals=('nan','none',''))
                 or _nv(p.get('구분',''), blank_vals=('nan','none',''))
                 or '차변')
        if vtype not in ('차변','대변','both'): vtype = '차변'
        acct_dir[acct] = vtype
    if not acct_dir: acct_dir = {'접대비': '차변'}

    if COL_ACCOUNT not in df.columns or COL_CLIENT not in df.columns: return {}
    if '구분' not in df.columns: return {}

    df_w = df.copy()
    gc = _get_gubun_col(df_w)
    if gc: df_w[gc] = df_w[gc].astype(str).str.strip()

    # 계정별로 개별 방향 적용 후 합치기
    parts = []
    for acct, vtype in acct_dir.items():
        sub = df_w[_account_match_flexible(df_w[COL_ACCOUNT], acct)].copy()
        if sub.empty: continue
        if vtype == '차변':
            sub['_amt'] = pd.to_numeric(sub[COL_DEBIT], errors='coerce').fillna(0)
        elif vtype == '대변':
            sub['_amt'] = pd.to_numeric(sub[COL_CREDIT], errors='coerce').fillna(0)
        else:
            sub['_amt'] = (pd.to_numeric(sub[COL_DEBIT], errors='coerce').fillna(0)
                          + pd.to_numeric(sub[COL_CREDIT], errors='coerce').fillna(0))
        parts.append(sub)
    if not parts: return {}
    filtered = pd.concat(parts, ignore_index=True)

    pivot = filtered.pivot_table(index=[COL_ACCOUNT, COL_CLIENT], columns='구분',
                                 values='_amt', aggfunc=['sum','count'], fill_value=0)
    if pivot.empty: return {}

    result = pd.DataFrame(index=pivot.index)
    for col in ['전기금액','당기금액','전기전표수','당기전표수']: result[col] = 0
    for (agg_fn, gubun) in pivot.columns:
        g = str(gubun).strip()
        col_key = f'{g}금액' if agg_fn == 'sum' else f'{g}전표수'
        result[col_key] = pivot[(agg_fn, gubun)].reindex(result.index).fillna(0)
    result = result.fillna(0)
    for c in ['전기전표수','당기전표수']:
        if c in result.columns: result[c] = result[c].astype(int)
    result['증감금액'] = result.get('당기금액', 0) - result.get('전기금액', 0)
    result['증감비율(%)'] = result.apply(
        lambda r: (r['증감금액']/r['전기금액']*100) if r.get('전기금액',0) != 0 else 0.0, axis=1)
    result = (result.assign(_abs=result['증감금액'].abs())
                    .sort_values([COL_ACCOUNT,'_abs','당기금액'], ascending=[True,False,False])
                    .drop(columns=['_abs']).reset_index())
    cols = list(result.columns)
    renames = {}
    if cols[0] != '계정명': renames[cols[0]] = '계정명'
    if len(cols)>1 and cols[1] != '거래처명': renames[cols[1]] = '거래처명'
    if renames: result = result.rename(columns=renames)

    out = {}
    if '계정명' in result.columns:
        for acct in result['계정명'].unique():
            sub   = result[result['계정명'] == acct].drop(columns=['계정명'])
            sname = _safe_sheet(f'비교_{re.sub(r"[^가-힣a-zA-Z0-9]","",str(acct))[:20]}')
            out[sname] = sub
    else:
        out['거래처_전기당기비교'] = result
    return out


# ── 3. 벤포드 분석 ────────────────────────────────────────────────────────────
def analyze_benford(df: pd.DataFrame, params_list: list) -> dict:
    targets = []
    for p in params_list:
        acct = _nv(p.get('계정과목',''))
        col  = _nv(p.get('금액열',''), blank_vals=('nan','none','')) or '차변'
        if col not in ('차변','대변'): col = '차변'
        if acct: targets.append((acct, col))
    if not targets: targets = list(DEFAULT_BENFORD_TARGETS)

    out, images = {}, []
    for acct, direction in targets:
        tcol   = COL_DEBIT if direction == '차변' else COL_CREDIT
        mask   = _account_match_flexible(df[COL_ACCOUNT], acct)
        subset = df[mask & (df[tcol] > 0)].copy()
        n      = len(subset)
        sheet_key = f'벤포드_{acct}_{direction}'
        if n < BENFORD_MIN_ROWS:
            out[sheet_key] = pd.DataFrame([{'계정':acct,'방향':direction,'숫자':'-','발생건수':0,
                                            '실제비율(%)':0,'이론비율(%)':0,'차이(%p)':0,
                                            '비고':f'데이터 부족({n}건)'}])
            continue
        subset['Digit'] = subset[tcol].apply(get_first_digit)
        dg     = subset[subset['Digit'] >= 1]['Digit']
        counts = dg.value_counts(normalize=True).sort_index()
        raw    = dg.value_counts().sort_index()
        img    = draw_benford_chart(acct, direction, counts, BENFORD_PROBS)
        if img: images.append((acct, direction, img))
        rows = []
        for d in range(1, 10):
            actual = counts.get(d, 0.0)
            theory = BENFORD_PROBS[d]
            rows.append({'계정':acct,'방향':direction,'숫자':d,'발생건수':int(raw.get(d,0)),
                         '실제비율(%)':round(actual*100,2),'이론비율(%)':round(theory*100,2),
                         '차이(%p)':round((actual-theory)*100,2),
                         '이상여부':'Y' if abs(actual-theory)>0.05 else ''})
        out[sheet_key] = pd.DataFrame(rows)

    if images: out['_benford_images'] = images   # 특수 키: save_results에서 차트 삽입
    return out


# ── 4. 데이터 개요 ────────────────────────────────────────────────────────────
def analyze_data_overview(df: pd.DataFrame, params_list: list) -> dict:
    summary = pd.DataFrame({
        '항목': ['총 행수','총 차변','총 대변','계정 수','시작일','종료일'],
        '값': [len(df),
               df[COL_DEBIT].sum()  if COL_DEBIT   in df.columns else 0,
               df[COL_CREDIT].sum() if COL_CREDIT  in df.columns else 0,
               df[COL_ACCOUNT].nunique() if COL_ACCOUNT in df.columns else 0,
               df[COL_DATE].min() if COL_DATE in df.columns else '-',
               df[COL_DATE].max() if COL_DATE in df.columns else '-']
    })
    if COL_ACCOUNT in df.columns and COL_DEBIT in df.columns and COL_CREDIT in df.columns:
        d = df[df[COL_DEBIT]!=0].groupby(COL_ACCOUNT)[COL_DEBIT].agg(['count','sum','mean','std'])
        d.columns = ['전표건수(차)','차변합계','평균금액(차)','표준편차(차)']
        c = df[df[COL_CREDIT]!=0].groupby(COL_ACCOUNT)[COL_CREDIT].agg(['count','sum','mean','std'])
        c.columns = ['전표건수(대)','대변합계','평균금액(대)','표준편차(대)']
        stats = pd.concat([d, c], axis=1).fillna(0).sort_values('차변합계',ascending=False).reset_index()
    else:
        stats = pd.DataFrame()
    return {'데이터개요_요약': summary, '데이터개요_계정별': stats}


# ── 5. 계정명 리스트 ──────────────────────────────────────────────────────────
def analyze_account_list(df: pd.DataFrame, params_list: list) -> dict:
    if COL_ACCOUNT not in df.columns:
        return {'계정명리스트': pd.DataFrame({'오류':['계정명 컬럼 없음']})}
    accs   = sorted({str(a).strip() for a in df[COL_ACCOUNT].dropna() if str(a).strip()})
    simple = pd.DataFrame({'계정명': accs})
    agg_map = {COL_DEBIT:['sum','count'], COL_CREDIT:['sum','count']}
    if COL_JOURNAL_ID in df.columns: agg_map[COL_JOURNAL_ID] = 'nunique'
    stats = df.groupby(COL_ACCOUNT).agg(agg_map).reset_index()
    base_cols = ['계정명','차변합계','차변건수','대변합계','대변건수']
    stats.columns = base_cols + (['전표개수'] if COL_JOURNAL_ID in df.columns else [])
    stats['최대금액'] = stats[['차변합계','대변합계']].max(axis=1)
    stats = stats.sort_values('최대금액', ascending=False).drop(columns=['최대금액'])
    return {'계정명_간단': simple, '계정명_통계': stats}


# ── 6. 사원별 집계 ────────────────────────────────────────────────────────────
def analyze_employee_summary(df: pd.DataFrame, params_list: list) -> pd.DataFrame:
    emp_col = next(
        (c for c in df.columns if any(n in str(c) for n in ['사원명','작성자','사용자','User','Employee'])),
        None)
    if emp_col is None:
        return pd.DataFrame({'오류':['사원명 컬럼 없음. 컬럼 목록: ' + str(list(df.columns))]})
    df_emp = df[df[emp_col].notna() & (df[emp_col].astype(str).str.strip() != '')].copy()
    if df_emp.empty: return pd.DataFrame({'오류':['사원명 데이터 없음']})

    jid = COL_JOURNAL_ID if COL_JOURNAL_ID in df.columns else None
    rows = []
    for emp in sorted(df_emp[emp_col].astype(str).unique()):
        sub   = df_emp[df_emp[emp_col].astype(str) == emp]
        d_sub = sub[sub[COL_DEBIT]  > 0] if COL_DEBIT  in sub.columns else pd.DataFrame()
        c_sub = sub[sub[COL_CREDIT] > 0] if COL_CREDIT in sub.columns else pd.DataFrame()
        dagg  = {COL_DEBIT:'sum'}
        cagg  = {COL_CREDIT:'sum'}
        if jid: dagg[jid]='nunique'; cagg[jid]='nunique'
        dg = d_sub.groupby(COL_ACCOUNT).agg(dagg).reset_index() if not d_sub.empty else pd.DataFrame()
        cg = c_sub.groupby(COL_ACCOUNT).agg(cagg).reset_index() if not c_sub.empty else pd.DataFrame()
        for i in range(max(len(dg), len(cg), 1)):
            row = {'사원명': emp if i == 0 else ''}
            if i < len(dg):
                row['차변계정명'] = dg.iloc[i][COL_ACCOUNT]
                row['차변금액']   = dg.iloc[i][COL_DEBIT]
                row['차변전표수'] = int(dg.iloc[i][jid]) if jid and jid in dg.columns else ''
            else:
                row.update({'차변계정명':'','차변금액':0,'차변전표수':''})
            if i < len(cg):
                row['대변계정명'] = cg.iloc[i][COL_ACCOUNT]
                row['대변금액']   = cg.iloc[i][COL_CREDIT]
                row['대변전표수'] = int(cg.iloc[i][jid]) if jid and jid in cg.columns else ''
            else:
                row.update({'대변계정명':'','대변금액':0,'대변전표수':''})
            rows.append(row)
    return pd.DataFrame(rows)


# ── 7. 일자차이 분석 ──────────────────────────────────────────────────────────
def analyze_date_difference(df: pd.DataFrame, params_list: list) -> dict:
    days_threshold = None
    for p in params_list:
        v = p.get('기준일수', p.get('일수'))
        if v is not None:
            try: days_threshold = int(float(v)); break
            except (TypeError, ValueError): pass
    if not days_threshold or days_threshold <= 0:
        return {'일자차이분석': pd.DataFrame({'안내':['기준일수 파라미터가 없거나 0 이하입니다.']})}

    reg_names = ['등록일자','등록일','작성일자','작성일','생성일자','입력일자','입력일']
    reg_col = next((c for c in df.columns if any(n in str(c) for n in reg_names)), None)
    if reg_col is None:
        return {'일자차이분석': pd.DataFrame({'오류':['등록일자 컬럼 없음',
                                                       f'컬럼: {list(df.columns[:15])}']})}

    df2 = df[(df[COL_DATE].notna()) & (df[reg_col].notna())].copy()
    if df2.empty:
        return {'일자차이분석': pd.DataFrame({'오류':['전표일자 또는 등록일자 없음']})}

    df2['일자차이'] = (df2[reg_col] - df2[COL_DATE]).dt.days
    filtered = df2[df2['일자차이'] >= days_threshold].copy()
    if filtered.empty:
        return {'일자차이분석': pd.DataFrame({'결과':[f'{days_threshold}일 이상 차이 전표 없음']})}

    j_sum = filtered.groupby(COL_JOURNAL_ID).agg({
        COL_DATE:'first', reg_col:'first', '일자차이':'first',
        COL_ACCOUNT: lambda x: ', '.join(x.unique()[:5]),
        COL_DEBIT:'sum', COL_CREDIT:'sum'
    }).reset_index().sort_values('일자차이', ascending=False)
    j_sum.columns = ['전표번호','전표일자','등록일자','일자차이','계정명','차변합계','대변합계']

    detail = filtered.sort_values(['일자차이', COL_JOURNAL_ID], ascending=[False, True])
    dc = [c for c in ['구분', COL_JOURNAL_ID, COL_DATE, reg_col, '일자차이',
                       COL_ACCOUNT, COL_DEBIT, COL_CREDIT, COL_CLIENT, COL_DESC] if c in detail.columns]
    return {'일자차이_요약': j_sum, '일자차이_상세': detail[dc]}


# ── 8. 상대계정 분석 ──────────────────────────────────────────────────────────
def analyze_counterpart(df: pd.DataFrame, params_list: list) -> dict:
    out = {}
    for p in params_list:
        acct      = _nv(p.get('계정과목',''))
        direction = (_nv(p.get('금액열',''), blank_vals=('nan','none',''))
                     or _nv(p.get('구분',''), blank_vals=('nan','none',''))
                     or '차변')
        if direction not in ('차변','대변'): direction = '차변'
        if not acct: continue
        tcol  = COL_DEBIT if direction == '차변' else COL_CREDIT
        mask  = _account_match_flexible(df[COL_ACCOUNT], acct) & (df[tcol] > 0)
        target = df[mask]
        if target.empty: continue
        jids    = target[COL_JOURNAL_ID].unique()
        related = df[df[COL_JOURNAL_ID].isin(jids)].copy()
        # 지정 방향의 반대 측(상대계정) 금액만 집계
        counter_col = COL_CREDIT if direction == '차변' else COL_DEBIT
        sum_label   = '대변합계' if direction == '차변' else '차변합계'
        cnt_label   = '대변건수' if direction == '차변' else '차변건수'
        summary = (related[related[counter_col] > 0]
                   .groupby(COL_ACCOUNT)[[counter_col]]
                   .agg(['sum','count']).reset_index())
        summary.columns = ['상대계정명', sum_label, cnt_label]
        summary = summary.sort_values(sum_label, ascending=False)
        sname   = _safe_sheet(f'상대_{re.sub(r"[^가-힣a-zA-Z0-9]","",acct)[:18]}')
        out[sname] = summary
    return out or {'상대계정분석': pd.DataFrame({'안내':['파라미터에 계정과목이 없습니다.']})}


# ── 9. 키워드 검색 ────────────────────────────────────────────────────────────
def analyze_keyword_search(df: pd.DataFrame, params_list: list) -> pd.DataFrame:
    keywords = [_nv(p.get('키워드','')) for p in params_list if _nv(p.get('키워드',''))]
    if not keywords: keywords = list(DEFAULT_KEYWORDS)
    if COL_DESC not in df.columns:
        return pd.DataFrame({'오류':['적요 컬럼 없음']})
    pattern = '|'.join(re.escape(k) for k in keywords)
    result  = df[df[COL_DESC].str.contains(pattern, na=False, regex=True)].copy()
    if result.empty:
        return pd.DataFrame({'결과':[f'검색어 없음: {", ".join(keywords)}']})
    result['AbsAmt'] = result[[COL_DEBIT, COL_CREDIT]].abs().max(axis=1)
    return result.sort_values('AbsAmt', ascending=False).drop(columns=['AbsAmt'])


# ── 10. 라운드넘버 분석 ───────────────────────────────────────────────────────
def analyze_round_numbers(df: pd.DataFrame, params_list: list) -> pd.DataFrame:
    ALL_UNITS = [100_000, 500_000, 1_000_000, 5_000_000, 10_000_000]
    records = []
    for p in params_list:
        acct     = _nv(p.get('계정과목',''))
        col_f    = _nv(p.get('금액열',''), blank_vals=('nan','none','')) or '차변'
        min_unit = _safe_float(p.get('최소금액'), 100_000.0)
        units    = [u for u in ALL_UNITS if u >= min_unit]
        sub      = df.copy()
        if acct and COL_ACCOUNT in sub.columns:
            sub = sub[sub[COL_ACCOUNT].astype(str).str.contains(acct, na=False)]
        amt_cols = [col_f] if col_f in ('차변','대변') and col_f in sub.columns \
                   else [c for c in ('차변','대변') if c in sub.columns]
        for col in amt_cols:
            work = sub[sub[col] > 0].copy()
            work['라운드단위'] = work[col].apply(
                lambda x: next((f'{u:,}원 배수' for u in sorted(units,reverse=True) if x%u==0), None))
            records.append(work[work['라운드단위'].notna()].assign(금액열=col))
    if not records: return pd.DataFrame({'결과':['라운드넘버 없음']})
    result  = pd.concat(records, ignore_index=True)
    amt_col = next((c for c in ('차변','대변','금액') if c in result.columns), None)
    out_cols = [c for c in ['구분', COL_DATE, COL_JOURNAL_ID, COL_ACCOUNT, '금액열', amt_col,
                             COL_DESC, COL_CLIENT, '라운드단위'] if c and c in result.columns]
    return result[out_cols].sort_values(amt_col, ascending=False) if amt_col else result[out_cols]


# ── 11. 특수관계자 분석 ───────────────────────────────────────────────────────
def analyze_related_party(df: pd.DataFrame, params_list: list) -> dict:
    parties = [_nv(p.get('거래처명','')) for p in params_list if _nv(p.get('거래처명',''))]
    if not parties:
        return {'특수관계자': pd.DataFrame({'안내':['task_list 특수관계자 시트에 거래처명을 입력하세요.']})}
    if COL_CLIENT not in df.columns:
        return {'특수관계자': pd.DataFrame({'오류':['거래처명 컬럼 없음']})}
    pattern = '|'.join(_related_party_pattern(p) for p in parties)
    related = df[df[COL_CLIENT].str.contains(pattern, na=False, regex=True)].copy()
    if related.empty:
        return {'특수관계자': pd.DataFrame({'결과':['해당 특수관계자 거래 없음']})}
    piv_d = related.pivot_table(index=COL_ACCOUNT, columns=COL_CLIENT, values=COL_DEBIT,
                                 aggfunc='sum', fill_value=0).reset_index()
    piv_c = related.pivot_table(index=COL_ACCOUNT, columns=COL_CLIENT, values=COL_CREDIT,
                                 aggfunc='sum', fill_value=0).reset_index()
    summary = related.groupby([COL_CLIENT, COL_ACCOUNT])[[COL_DEBIT, COL_CREDIT]]\
                     .agg(['sum','count']).reset_index()
    summary.columns = ['거래처명','계정명','차변합계','차변건수','대변합계','대변건수']
    return {'특수관계자_차변피벗': piv_d,
            '특수관계자_대변피벗': piv_c,
            '특수관계자_요약':     summary,
            '특수관계자_상세':     related}


# ── 12. 자산 vs 부채 교차 ────────────────────────────────────────────────────
def analyze_asset_liability_cross(df: pd.DataFrame, params_list: list) -> pd.DataFrame:
    assets = [_nv(p.get('계정과목','')) for p in params_list
              if str(p.get('구분','')).strip() == '자산' and _nv(p.get('계정과목',''))]
    liabs  = [_nv(p.get('계정과목','')) for p in params_list
              if str(p.get('구분','')).strip() == '부채' and _nv(p.get('계정과목',''))]
    if not assets or not liabs:
        return pd.DataFrame({'안내':['task_list 자산부채교차 시트에 구분(자산/부채)·계정과목을 입력하세요.']})
    am = pd.Series(False, index=df.index)
    for a in assets: am |= df[COL_ACCOUNT].str.contains(a, na=False, regex=False)
    lm = pd.Series(False, index=df.index)
    for l in liabs:  lm |= df[COL_ACCOUNT].str.contains(l, na=False, regex=False)
    adf = df[am & (df[COL_DEBIT]  > 0)]
    ldf = df[lm & (df[COL_CREDIT] > 0)]
    if adf.empty or ldf.empty: return pd.DataFrame({'결과':['교차 데이터 없음']})
    ga = adf.groupby(COL_CLIENT).agg({COL_DEBIT:'sum',  COL_ACCOUNT: lambda x: ','.join(set(x))}).reset_index()
    gl = ldf.groupby(COL_CLIENT).agg({COL_CREDIT:'sum', COL_ACCOUNT: lambda x: ','.join(set(x))}).reset_index()
    ga.columns = ['거래처명','자산_금액','자산_계정']
    gl.columns = ['거래처명','부채_금액','부채_계정']
    merged = pd.merge(ga, gl, on='거래처명', how='inner').sort_values('자산_금액', ascending=False)
    return merged if not merged.empty else pd.DataFrame({'결과':['동시 발생 거래처 없음']})


# ── 13. 매출 vs 비용 교차 ────────────────────────────────────────────────────
def analyze_revenue_expense_cross(df: pd.DataFrame, params_list: list) -> pd.DataFrame:
    revs = [_nv(p.get('계정과목','')) for p in params_list
            if str(p.get('구분','')).strip() == '매출' and _nv(p.get('계정과목',''))]
    exps = [_nv(p.get('계정과목','')) for p in params_list
            if str(p.get('구분','')).strip() == '비용' and _nv(p.get('계정과목',''))]
    if not revs or not exps:
        return pd.DataFrame({'안내':['task_list 매출비용교차 시트에 구분(매출/비용)·계정과목을 입력하세요.']})
    rm = pd.Series(False, index=df.index)
    for r in revs: rm |= df[COL_ACCOUNT].str.contains(r, na=False, regex=False)
    em = pd.Series(False, index=df.index)
    for e in exps: em |= df[COL_ACCOUNT].str.contains(e, na=False, regex=False)
    rdf = df[rm & (df[COL_CREDIT] > 0)]
    edf = df[em & (df[COL_DEBIT]  > 0)]
    if rdf.empty or edf.empty: return pd.DataFrame({'결과':['매출 또는 비용 데이터 없음']})
    gr = rdf.groupby(COL_CLIENT).agg({COL_CREDIT:'sum', COL_ACCOUNT: lambda x: ','.join(set(x))}).reset_index()
    ge = edf.groupby(COL_CLIENT).agg({COL_DEBIT:'sum',  COL_ACCOUNT: lambda x: ','.join(set(x))}).reset_index()
    gr.columns = ['거래처명','매출_금액','매출_계정']
    ge.columns = ['거래처명','비용_금액','비용_계정']
    merged = pd.merge(gr, ge, on='거래처명', how='inner').sort_values('매출_금액', ascending=False)
    return merged if not merged.empty else pd.DataFrame({'결과':['동시 발생 거래처 없음']})


# ── 14. 심층분석 (계정별 Top) ─────────────────────────────────────────────────
def analyze_top_accounts(df: pd.DataFrame, params_list: list) -> dict:
    config = []
    for p in params_list:
        acct = _nv(p.get('계정과목',''))
        if not acct: continue
        top_n = int(_safe_float(p.get('개수', p.get('top_n', 10)), 10))
        direction = _nv(p.get('금액열',''), blank_vals=('nan','none','')) or 'both'
        if direction not in ('차변','대변','both'): direction = 'both'
        config.append((acct, top_n, direction))
    if not config: return {'심층분석': pd.DataFrame({'안내':['계정과목 파라미터가 없습니다.']})}

    gc = _get_gubun_col(df)
    has_g = gc is not None
    out = {}

    for idx, (acct_name, top_n, direction) in enumerate(config, 1):
        filtered = df[df[COL_ACCOUNT].str.contains(acct_name, na=False, regex=False)]
        if filtered.empty: continue
        sname     = _safe_sheet(f'Top_{idx}_{re.sub(r"[^가-힣a-zA-Z0-9]","",acct_name)[:16]}')
        grp_cols  = [gc, COL_CLIENT] if has_g else [COL_CLIENT]

        debit_top = pd.DataFrame()
        if direction in ('차변', 'both'):
            d_rows = filtered[filtered[COL_DEBIT] > 0]
            if not d_rows.empty and COL_CLIENT in df.columns:
                debit_top = (d_rows.groupby(grp_cols)[COL_DEBIT].agg(['count','sum'])
                                   .reset_index().sort_values('sum', ascending=False)
                                   .head(top_n*(2 if has_g else 1)))
                rn = {grp_cols[-1]:'거래처명','count':'전표수(차)','sum':'차변금액'}
                if has_g: rn[gc] = '구분'
                debit_top = debit_top.rename(columns=rn)
                debit_top.insert(0, '계정명', acct_name)

        credit_top = pd.DataFrame()
        if direction in ('대변', 'both'):
            c_rows = filtered[filtered[COL_CREDIT] > 0]
            if not c_rows.empty and COL_CLIENT in df.columns:
                credit_top = (c_rows.groupby(grp_cols)[COL_CREDIT].agg(['count','sum'])
                                    .reset_index().sort_values('sum', ascending=False)
                                    .head(top_n*(2 if has_g else 1)))
                rn = {grp_cols[-1]:'거래처명','count':'전표수(대)','sum':'대변금액'}
                if has_g: rn[gc] = '구분'
                credit_top = credit_top.rename(columns=rn)
                credit_top.insert(0, '계정명', acct_name)

        if direction == 'both':
            combined = pd.concat([debit_top, credit_top], axis=1)
        elif direction == '차변':
            combined = debit_top
        else:
            combined = credit_top

        if not combined.empty:
            out[sname] = combined

    return out or {'심층분석': pd.DataFrame({'결과':['데이터 없음']})}


# ── 15. AI 계정별 분석 ────────────────────────────────────────────────────────
def analyze_ai_preparation(df: pd.DataFrame, params_list: list) -> dict:
    targets = [_nv(p.get('계정과목','')) for p in params_list if _nv(p.get('계정과목',''))]
    if not targets: return {'AI분석': pd.DataFrame({'안내':['계정과목 파라미터 없음']})}
    out = {}
    for acct in targets:
        filtered = df[df[COL_ACCOUNT].str.contains(acct, na=False, regex=False)].copy()
        if filtered.empty: continue
        safe_nm   = re.sub(r'[\\/*?:\[\]]', '', acct)[:10]

        filtered['YM'] = pd.to_datetime(filtered[COL_DATE], errors='coerce').dt.strftime('%Y-%m')
        monthly = filtered.groupby('YM')[[COL_DEBIT, COL_CREDIT]].agg(['sum','count']).reset_index()
        monthly.columns = ['YM','차변합계','차변건수','대변합계','대변건수']

        filtered['MaxAmt'] = filtered[[COL_DEBIT, COL_CREDIT]].max(axis=1)
        sample = (filtered[filtered['MaxAmt'] >= filtered['MaxAmt'].quantile(0.90)].copy()
                  if len(filtered) > 10 else filtered.copy())
        if COL_CLIENT in sample.columns:
            sample[COL_CLIENT] = sample[COL_CLIENT].apply(get_safe_client_name)
        for col in MASK_TARGET_COLS:
            if col in sample.columns: sample[col] = sample[col].apply(mask_sensitive_info)
        sample = sample.drop(columns=['MaxAmt','YM'], errors='ignore')

        out[_safe_sheet(f'AI_{safe_nm}_월별')] = monthly
        out[_safe_sheet(f'AI_{safe_nm}_샘플')] = sample

    if GLOBAL_SAFE_MAP:
        out['_암호해독표'] = pd.DataFrame(list(GLOBAL_SAFE_MAP.items()), columns=['실명','가명'])
    return out or {'AI분석': pd.DataFrame({'결과':['분석 대상 없음']})}


# ── 16. 데이터·헤더 확인 ──────────────────────────────────────────────────────
def analyze_header_check(df: pd.DataFrame, params_list: list) -> pd.DataFrame:
    required = [(COL_DEBIT,'차변'),(COL_CREDIT,'대변'),(COL_ACCOUNT,'계정명'),
                (COL_CLIENT,'거래처명'),(COL_DATE,'전표일자'),(COL_JOURNAL_ID,'전표번호')]
    rows = [{'항목': name, '인식': 'O' if col in df.columns else 'X',
              '실제컬럼': col if col in df.columns else '없음'} for col, name in required]
    gc = _get_gubun_col(df)
    if gc:
        rows.append({'항목':'구분(당기/전기)','인식':'O','실제컬럼':str(df[gc].unique().tolist())})
    rows.append({'항목':'총 행수','인식':'-','실제컬럼':str(len(df))})
    rows.append({'항목':'전체 컬럼','인식':'-','실제컬럼':str(list(df.columns))})
    return pd.DataFrame(rows)


# ── 17. 거래처 분석 ───────────────────────────────────────────────────────────
def analyze_client_detail(df: pd.DataFrame, params_list: list) -> dict:
    out = {}
    for i, p in enumerate(params_list, 1):
        accts    = [a.strip() for a in str(p.get('계정과목','')).split(',')
                    if a.strip() and a.strip().lower() not in ('nan','(전체)','')]
        clients  = [c.strip() for c in str(p.get('거래처명','')).split(',')
                    if c.strip() and c.strip().lower() not in ('nan','')]
        vtype    = _nv(p.get('금액열',''), blank_vals=('nan','none','')) or 'both'
        if vtype not in ('차변','대변','both'): vtype = 'both'
        job_name = _nv(p.get('작업명','')) or f'거래처{i:02d}'
        if not clients: continue

        mask_a = pd.Series(True, index=df.index) if not accts else pd.Series(False, index=df.index)
        for acct in accts: mask_a |= _account_match_flexible(df[COL_ACCOUNT], acct)
        mask_c = pd.Series(False, index=df.index)
        cli_col = df[COL_CLIENT].fillna('').astype(str)
        for cli in clients: mask_c |= cli_col.str.contains(cli, na=False, regex=False)

        filtered = df[mask_a & mask_c].copy()
        if filtered.empty: continue
        if vtype == '차변': filtered = filtered[filtered[COL_DEBIT]  > 0]
        elif vtype == '대변': filtered = filtered[filtered[COL_CREDIT] > 0]
        if filtered.empty: continue

        filtered['YM'] = pd.to_datetime(filtered[COL_DATE], errors='coerce').dt.strftime('%Y-%m')
        monthly = filtered.groupby('YM').agg({COL_DEBIT:'sum',COL_CREDIT:'sum'}).reset_index()
        monthly.columns = ['YM','차변합계','대변합계']
        monthly['합계'] = monthly['차변합계'] + monthly['대변합계']

        dc = [c for c in ['구분', COL_DATE, COL_JOURNAL_ID, COL_ACCOUNT,
                           COL_DEBIT, COL_CREDIT, COL_CLIENT, COL_DESC] if c in filtered.columns]
        sname = _safe_sheet(f'거래처_{re.sub(r"[^가-힣a-zA-Z0-9]","",job_name)[:18]}')
        out[sname]               = filtered[dc]
        out[sname + '_월별합산'] = monthly

    return out or {'거래처분석': pd.DataFrame({'안내':['파라미터에 거래처명이 없습니다.']})}


# ── 18. 벤포드 이탈 상세 추출 ────────────────────────────────────────────────
def analyze_benford_deviation(df: pd.DataFrame, params_list: list) -> dict:
    targets, threshold, max_per = [], 0.03, 500
    for p in params_list:
        acct = _nv(p.get('계정과목',''))
        col  = _nv(p.get('금액열',''), blank_vals=('nan','none','')) or '차변'
        if col not in ('차변','대변'): col = '차변'
        if acct: targets.append((acct, col))
        t = _safe_float(p.get('임계값'), 0.03)
        if 0.001 <= t <= 0.3: threshold = t
        m = p.get('최대건수')
        if m:
            try: max_per = int(m)
            except Exception: pass
    if not targets: targets = list(DEFAULT_BENFORD_TARGETS)

    all_detail, summary_rows = [], []
    for acct, direction in targets:
        tcol   = COL_DEBIT if direction == '차변' else COL_CREDIT
        mask   = _account_match_flexible(df[COL_ACCOUNT], acct)
        subset = df[mask & (df[tcol] > 0)].copy()
        n      = len(subset)
        if n < BENFORD_MIN_ROWS:
            summary_rows.append({'계정':acct,'방향':direction,'선행자릿수':'-','이탈정도':'-',
                                  '추출건수':'-','전체건수':n,'비고':'데이터 부족'})
            continue
        subset['Digit'] = subset[tcol].apply(get_first_digit)
        dg     = subset[subset['Digit'] >= 1]['Digit']
        counts = dg.value_counts(normalize=True).sort_index()
        deviant_set = {d for d in range(1,10) if abs(counts.get(d,0)-BENFORD_PROBS[d]) >= threshold}
        if not deviant_set:
            summary_rows.append({'계정':acct,'방향':direction,'선행자릿수':'-','이탈정도':'-',
                                  '추출건수':'-','전체건수':n,'비고':'임계값 이상 이탈 없음'})
            continue
        deviant_df = subset[subset['Digit'].isin(deviant_set)].copy()
        deviant_df['이탈정도'] = deviant_df['Digit'].map(
            lambda d: round(counts.get(d,0)-BENFORD_PROBS[d],3))
        out_cols = [c for c in ['구분', COL_DATE, COL_JOURNAL_ID, COL_ACCOUNT, COL_CLIENT, COL_DESC,
                                 tcol, 'Digit', '이탈정도'] if c in deviant_df.columns]
        for d in sorted(deviant_set):
            sub = deviant_df[deviant_df['Digit'] == d].head(max_per)
            summary_rows.append({'계정':acct,'방향':direction,'선행자릿수':d,
                                  '이탈정도':round(counts.get(d,0)-BENFORD_PROBS[d],3),
                                  '추출건수':len(sub),
                                  '전체건수':int((deviant_df['Digit']==d).sum()),'비고':''})
            exp = sub[[c for c in out_cols if c in sub.columns]].copy()
            exp['계정분류'] = f'{acct}({direction})'
            all_detail.append(exp)

    out = {}
    if summary_rows: out['벤포드이탈_요약'] = pd.DataFrame(summary_rows)
    if all_detail:   out['벤포드이탈_상세'] = pd.concat(all_detail, ignore_index=True)
    return out or {'벤포드이탈': pd.DataFrame({'결과':['추출 데이터 없음']})}


# ── 19. 월별 전계정 분석 ──────────────────────────────────────────────────────
def analyze_monthly_full_account(df: pd.DataFrame, params_list: list) -> pd.DataFrame:
    work = df.copy()
    work['Month'] = pd.to_datetime(work[COL_DATE], errors='coerce').dt.month
    grp = ['구분', COL_ACCOUNT, 'Month'] if '구분' in work.columns else [COL_ACCOUNT, 'Month']
    return work.groupby(grp)[[COL_DEBIT, COL_CREDIT]].sum().reset_index()


# ── 20. 당기증감분석 (계정별 거래처별) ───────────────────────────────────────
def analyze_balance_movement(df: pd.DataFrame, params_list: list) -> dict:
    """전기 계정별_거래처별명세에서 기초잔액, 당기 분개장에서 증감 산출하여 기말잔액 계산.
    자산(차변): 기초잔액 + 당기증가(차변) - 당기감소(대변) = 기말잔액
    부채(대변): 기초잔액 + 당기증가(대변) - 당기감소(차변) = 기말잔액
    """
    global _COMPANY_DIR
    if _COMPANY_DIR is None:
        return {'잔액증감분석': pd.DataFrame({'오류': ['company_dir 미설정']})}

    prev_dir = os.path.join(_COMPANY_DIR, 'data', 'previous')
    if not os.path.isdir(prev_dir):
        return {'잔액증감분석': pd.DataFrame({'오류': [f'previous 폴더 없음: {prev_dir}']})}

    # ── 전기 명세 파일 탐색 ──
    prev_file = None
    for f in sorted(os.listdir(prev_dir)):
        if ('계정별' in f or '거래처별' in f) and f.endswith('.xlsx') and not f.startswith('~$'):
            if '명세' in f or '거래처별' in f:
                prev_file = os.path.join(prev_dir, f)
                break
    if not prev_file:
        return {'잔액증감분석': pd.DataFrame({'오류': ['전기 계정별_거래처별명세 파일 없음']})}

    # ── 거래처 매핑 파일 탐색 ──
    vendor_mapping = {}
    for f in sorted(os.listdir(prev_dir)):
        if '매핑' in f and not f.startswith('~$'):
            map_path = os.path.join(prev_dir, f)
            try:
                if f.endswith('.csv'):
                    for enc in ['utf-8', 'cp949', 'euc-kr']:
                        try:
                            mdf = pd.read_csv(map_path, encoding=enc)
                            break
                        except UnicodeDecodeError:
                            continue
                else:
                    mdf = pd.read_excel(map_path)
                if '분개장거래처' in mdf.columns and '전기명세거래처' in mdf.columns:
                    for _, row in mdf.iterrows():
                        acct = str(row.get('계정명', '')).strip()
                        j_v = str(row.get('분개장거래처', '')).strip()
                        p_v = str(row.get('전기명세거래처', '')).strip()
                        if j_v and p_v and p_v not in ('nan', ''):
                            vendor_mapping[(acct, j_v)] = p_v
                    print(f'    └ 거래처 매핑: {len(vendor_mapping)}건')
            except Exception as e:
                print(f'    ⚠️ 매핑 파일 로드 실패: {e}')
            break

    # ── 파라미터 로드 (params_list → 직접 탐색 fallback) ──
    targets = []
    for p in params_list:
        acct = _nv(p.get('계정명', p.get('계정과목', '')))
        gubun = _nv(p.get('구분', ''))
        if acct and gubun in ('차변', '대변', '자산', '부채'):
            targets.append((acct, gubun))

    if not targets:
        task_path = os.path.join(
            _COMPANY_DIR, f'task_list_{os.path.basename(_COMPANY_DIR)}.xlsx')
        if os.path.isfile(task_path):
            txl = pd.ExcelFile(task_path)
            for s in txl.sheet_names:
                if '증감' in s and ('거래처' in s or '계정' in s or '게정' in s):
                    tdf = pd.read_excel(txl, sheet_name=s, header=None)
                    for i, row in tdf.iterrows():
                        vals = [str(v).strip() for v in row if pd.notna(v)]
                        if '구분' in vals:
                            for j in range(i + 1, len(tdf)):
                                drow = tdf.iloc[j]
                                a = str(drow.iloc[0]).strip() if pd.notna(drow.iloc[0]) else ''
                                g = str(drow.iloc[1]).strip() if len(drow) > 1 and pd.notna(drow.iloc[1]) else ''
                                if a and a not in ('nan', '') and g in ('차변', '대변', '자산', '부채'):
                                    targets.append((a, g))
                            break
                    break

    if not targets:
        return {'잔액증감분석': pd.DataFrame({'안내': ['파라미터에 계정명/구분이 없습니다.']})}

    print(f'    └ 분석 대상: {len(targets)}개 계정')

    # ── 전기 명세 ExcelFile ──
    prev_xl = pd.ExcelFile(prev_file)
    prev_sheets = prev_xl.sheet_names

    def _find_prev_sheet(acct_name):
        if acct_name in prev_sheets:
            return acct_name
        norm = acct_name.replace(' ', '')
        for s in prev_sheets:
            if s.replace(' ', '') == norm:
                return s
        norm2 = re.sub(r'[\s()（）]', '', acct_name)
        for s in prev_sheets:
            if re.sub(r'[\s()（）]', '', s) == norm2:
                return s
        for s in prev_sheets:
            if acct_name in s or s in acct_name:
                return s
        return None

    def _load_prev_balances(sheet_name):
        pdf = pd.read_excel(prev_xl, sheet_name=sheet_name, header=0)
        pdf.columns = [str(c).strip() for c in pdf.columns]
        vendor_col = next((c for c in pdf.columns if '거래처' in c), None)
        bal_col = next((c for c in pdf.columns if '잔' in c), None)
        if not vendor_col or not bal_col:
            return {}
        balances = {}
        for _, row in pdf.iterrows():
            vendor = str(row[vendor_col]).strip() if pd.notna(row[vendor_col]) else ''
            if not vendor or vendor.replace(' ', '') in ('합계:', '합계', 'nan', ''):
                continue
            try:
                bal = float(row[bal_col]) if pd.notna(row[bal_col]) else 0
            except (ValueError, TypeError):
                bal = 0
            balances[vendor] = balances.get(vendor, 0) + bal
        return balances

    # ── 계정별 분석 실행 ──
    all_results = {}

    for acct_name, gubun in targets:
        is_asset = gubun in ('차변', '자산')

        prev_sheet = _find_prev_sheet(acct_name)
        prev_balances = _load_prev_balances(prev_sheet) if prev_sheet else {}
        if prev_sheet:
            print(f'      {acct_name}: 전기 {len(prev_balances)}건 (시트: {prev_sheet})')
        else:
            print(f'      {acct_name}: 전기 시트 없음')

        mask = _account_match_flexible(df[COL_ACCOUNT], acct_name)
        current = df[mask].copy()

        journal_vendors = {}
        if not current.empty and COL_CLIENT in current.columns:
            for vendor, group in current.groupby(COL_CLIENT):
                v = str(vendor).strip()
                if v and v != 'nan':
                    journal_vendors[v] = {
                        'debit': group[COL_DEBIT].sum() if COL_DEBIT in group.columns else 0,
                        'credit': group[COL_CREDIT].sum() if COL_CREDIT in group.columns else 0,
                    }

        if not prev_balances and not journal_vendors:
            continue

        acct_map_j2p = {k[1]: v for k, v in vendor_mapping.items() if k[0] == acct_name}
        acct_map_p2j = {v: k[1] for k, v in vendor_mapping.items() if k[0] == acct_name}

        all_vendors = set(prev_balances.keys()) | set(journal_vendors.keys())
        rows = []
        matched_prev = set()
        matched_journal = set()

        for vendor in sorted(all_vendors):
            prev_bal = prev_balances.get(vendor, 0)
            prev_key = vendor
            if prev_bal == 0 and vendor in acct_map_j2p:
                prev_key = acct_map_j2p[vendor]
                prev_bal = prev_balances.get(prev_key, 0)

            jv = journal_vendors.get(vendor, None)
            journal_key = vendor
            if jv is None and vendor in acct_map_p2j:
                journal_key = acct_map_p2j[vendor]
                jv = journal_vendors.get(journal_key, None)
            if jv is None:
                jv = {'debit': 0, 'credit': 0}

            matched_prev.add(prev_key)
            matched_journal.add(journal_key)

            debit_sum = jv['debit']
            credit_sum = jv['credit']

            if is_asset:
                increase = debit_sum
                decrease = credit_sum
                ending = prev_bal + increase - decrease
                row = {
                    '계정명': acct_name, '거래처명': vendor,
                    '기초잔액': prev_bal, '당기증가': increase,
                    '당기감소': decrease, '기말잔액': ending,
                }
            else:
                increase = credit_sum
                decrease = debit_sum
                ending = prev_bal + increase - decrease
                row = {
                    '계정명': acct_name, '거래처명': vendor,
                    '기초잔액': prev_bal, '당기감소': decrease,
                    '당기증가': increase, '기말잔액': ending,
                }
            rows.append(row)

        if rows:
            result_df = pd.DataFrame(rows)
            amount_cols = [c for c in result_df.columns if c not in ('계정명', '거래처명')]
            result_df = result_df[result_df[amount_cols].abs().sum(axis=1) > 0]

            if not result_df.empty:
                result_df = result_df.sort_values('기말잔액', key=abs, ascending=False)
                total = {c: result_df[c].sum() for c in amount_cols}
                total['계정명'] = acct_name
                total['거래처명'] = '합  계'
                result_df = pd.concat(
                    [result_df, pd.DataFrame([total])], ignore_index=True)
                sname = _safe_sheet(
                    f'증감_{re.sub(r"[^가-힣a-zA-Z0-9]", "", acct_name)[:20]}')
                all_results[sname] = result_df

    return all_results or {'잔액증감분석': pd.DataFrame({'결과': ['분석 대상 없음']})}


# ── 21. 총계정원장 ────────────────────────────────────────────────────────────
def analyze_general_ledger(df: pd.DataFrame, params_list: list) -> dict:
    """계정별 월별 차변/대변 집계. 연도 2개 이상이면 연도 비교 형식(행=월, 열=연도)으로 출력.
    구분(자산/부채/매출액)은 시트명 접두어로 사용."""
    # 계정명 정규화를 루프 밖에서 1회만 계산 (52회 반복 방지)
    norm_col = df[COL_ACCOUNT].fillna('').astype(str).apply(_normalize_account_for_match)

    out = {}
    for p in params_list:
        acct  = _nv(p.get('계정과목', ''))
        gubun = _nv(p.get('구분', ''))
        if not acct: continue

        norm_user = _normalize_account_for_match(acct)
        mask = norm_col.str.startswith(norm_user)
        if not mask.any():
            mask = norm_col.str.contains(re.escape(norm_user), na=False, regex=True, case=False)
        if not mask.any():
            mask = norm_col.apply(lambda nv: norm_user in nv or nv in norm_user)

        subset = df[mask].copy()
        if subset.empty:
            continue

        subset['YM']    = pd.to_datetime(subset[COL_DATE], errors='coerce').dt.strftime('%Y-%m')
        subset          = subset[subset['YM'].notna()]
        subset['Year']  = subset['YM'].str[:4]
        subset['Month'] = subset['YM'].str[5:7].astype(int)

        years = sorted(subset['Year'].unique())
        dr    = subset[subset[COL_DEBIT]  > 0]
        cr    = subset[subset[COL_CREDIT] > 0]

        if len(years) >= 2:
            da = dr.groupby(['Year', 'Month']).agg(
                차변합계=(COL_DEBIT, 'sum'), 차변건수=(COL_DEBIT, 'count')).reset_index()
            ca = cr.groupby(['Year', 'Month']).agg(
                대변합계=(COL_CREDIT, 'sum'), 대변건수=(COL_CREDIT, 'count')).reset_index()
            ym = pd.merge(da, ca, on=['Year', 'Month'], how='outer').fillna(0)

            result = pd.DataFrame({'Month': range(1, 13)})
            for year in years:
                yd = ym[ym['Year'] == year][
                    ['Month', '차변합계', '대변합계', '차변건수', '대변건수']
                ].rename(columns={c: f'{year}_{c}' for c in ['차변합계', '대변합계', '차변건수', '대변건수']})
                result = pd.merge(result, yd, on='Month', how='left').fillna(0)
            for year in years:
                result[f'{year}_차변건수'] = result[f'{year}_차변건수'].astype(int)
                result[f'{year}_대변건수'] = result[f'{year}_대변건수'].astype(int)
            if len(years) == 2:
                y0, y1 = years[0], years[1]
                result['증감_차변합계'] = result[f'{y1}_차변합계'] - result[f'{y0}_차변합계']
                result['증감_대변합계'] = result[f'{y1}_대변합계'] - result[f'{y0}_대변합계']
            result.insert(0, '월', result['Month'].apply(lambda x: f'{x}월'))
            result = result.drop(columns=['Month'])
        else:
            da = dr.groupby('YM').agg(차변합계=(COL_DEBIT, 'sum'), 차변건수=(COL_DEBIT, 'count')).reset_index()
            ca = cr.groupby('YM').agg(대변합계=(COL_CREDIT, 'sum'), 대변건수=(COL_CREDIT, 'count')).reset_index()
            result = pd.merge(da, ca, on='YM', how='outer').fillna(0).sort_values('YM')
            result['차변건수'] = result['차변건수'].astype(int)
            result['대변건수'] = result['대변건수'].astype(int)
            result = result.rename(columns={'YM': '월'})

        total = {'월': '합  계'}
        for col in result.columns[1:]:
            total[col] = result[col].sum()
        result = pd.concat([result, pd.DataFrame([total])], ignore_index=True)

        prefix = f'{gubun}_' if gubun else ''
        sname  = _safe_sheet(f'총계정원장_{prefix}{re.sub(r"[^가-힣a-zA-Z0-9]", "", acct)[:18]}')
        out[sname] = result

    return out or {'총계정원장': pd.DataFrame({'결과': ['분석 대상 없음']})}


# ── 22. 은행조회서 완전성 ──────────────────────────────────────────────────────
_BANK_DEFAULT_ACCOUNTS = [
    '이자비용', '단기차입금', '장기차입금',
    '유동성장기부채', '유동성장기차입금',
    '전환사채', '유동성전환사채',
    '전환상환우선주', '유동성전환상환우선주',
]

_FI_PATTERN = re.compile(
    r'([\w가-힣()（）]*'
    r'(?:저축은행|저축|은행|금고|신협|조합|증권|캐피탈|카드|보험|파이낸스|크레딧|리스))'
)


def _extract_fi(client_name: str) -> str:
    name = str(client_name).strip()
    if not name or name.lower() in ('nan', 'none', ''):
        return ''
    m = _FI_PATTERN.search(name)
    return m.group(1) if m else ''


def analyze_bank_confirmation(df: pd.DataFrame, params_list: list) -> dict:
    """22. 은행조회서 완전성: 차입금/이자비용 등 관련 계정 상세내역 + 금융기관 요약."""
    # 파라미터 시트에 계정 목록이 있으면 사용, 없으면 기본값
    account_list = [
        _nv(p.get('계정과목', '') or p.get('계정명', '') or p.get('계정', ''))
        for p in params_list
    ]
    account_list = [a for a in account_list if a]
    if not account_list:
        account_list = list(_BANK_DEFAULT_ACCOUNTS)

    all_rows = []
    for acct in account_list:
        mask = df[COL_ACCOUNT].astype(str).str.contains(acct, na=False)
        sub = df[mask].copy()
        if sub.empty:
            continue
        sub.insert(0, '조회계정', acct)
        all_rows.append(sub)

    if not all_rows:
        return {'은행조회서완전성': pd.DataFrame({'안내': ['해당 계정의 전표 내역이 없습니다.']})}

    combined = pd.concat(all_rows, ignore_index=True)

    # 금융기관명 추출 (거래처명 바로 뒤에 컬럼 삽입)
    if COL_CLIENT in combined.columns:
        combined['금융기관명'] = combined[COL_CLIENT].apply(_extract_fi)
    else:
        combined['금융기관명'] = ''

    priority = ['조회계정']
    if '구분' in combined.columns: priority.append('구분')
    for c in [COL_DATE, COL_JOURNAL_ID, COL_ACCOUNT, COL_DEBIT, COL_CREDIT, COL_CLIENT]:
        if c in combined.columns: priority.append(c)
    priority.append('금융기관명')
    if COL_DESC in combined.columns: priority.append(COL_DESC)
    other = [c for c in combined.columns if c not in priority]
    combined = combined[priority + other]

    acct_order = {a: i for i, a in enumerate(account_list)}
    combined['_sort'] = combined['조회계정'].map(acct_order).fillna(999)
    sort_cols = ['_sort', COL_DATE] if COL_DATE in combined.columns else ['_sort']
    combined = combined.sort_values(sort_cols).drop(columns=['_sort'])

    results = {'은행조회서완전성': combined}

    # 금융기관별 요약 피벗
    has_fi = combined[combined['금융기관명'].astype(str).str.strip() != '']
    if not has_fi.empty and COL_CLIENT in has_fi.columns:
        pivot = has_fi.groupby(['금융기관명', '조회계정']).size().unstack(fill_value=0)
        acct_cols = [a for a in account_list if a in pivot.columns]
        pivot = pivot.reindex(columns=acct_cols, fill_value=0)
        try:
            pivot_mark = pivot.map(lambda x: '○' if x > 0 else '-')
        except AttributeError:
            pivot_mark = pivot.applymap(lambda x: '○' if x > 0 else '-')

        raw_clients = (
            has_fi.groupby('금융기관명')[COL_CLIENT]
            .apply(lambda s: ', '.join(sorted(s.dropna().astype(str).unique())))
            .rename('거래처명(원본)')
        )
        agg_kw = {'전표건수': ('조회계정', 'count')}
        if COL_DEBIT  in has_fi.columns: agg_kw['차변합계'] = (COL_DEBIT,  'sum')
        if COL_CREDIT in has_fi.columns: agg_kw['대변합계'] = (COL_CREDIT, 'sum')
        totals = has_fi.groupby('금융기관명').agg(**agg_kw)

        summary = pivot_mark.join(raw_clients).join(totals).reset_index()
        summary.insert(1, '조회서발송', 'Y')
        results['금융기관_조회서목록'] = summary

    return results



# ── 23. 계정별 상세거래내역 ──────────────────────────────────────────────────

def analyze_account_transaction_detail(df: pd.DataFrame, params_list: list) -> dict:
    """23. 계정별 상세거래내역: 계정과목별 건별 전표 내역 추출 (차변/대변/전체 선택)."""
    targets = []
    for p in params_list:
        acct = _nv(p.get('계정과목', p.get('계정명', '')))
        col_f = _nv(p.get('금액 유형', p.get('금액유형', p.get('차대구분', ''))),
                    blank_vals=('nan', 'none', ''))
        if acct:
            targets.append((acct, col_f))

    if not targets:
        return {'계정별상세내역': pd.DataFrame({'안내': ['파라미터에 계정과목이 없습니다.']})}

    all_results = {}
    for acct_name, col_f in targets:
        mask = _account_match_flexible(df[COL_ACCOUNT], acct_name)
        sub = df[mask].copy()

        if sub.empty:
            continue

        # 차변/대변 필터
        col_f_norm = col_f.replace(' ', '')
        if col_f_norm in ('차변', '차변만'):
            sub = sub[sub[COL_DEBIT] > 0]
        elif col_f_norm in ('대변', '대변만'):
            sub = sub[sub[COL_CREDIT] > 0]
        # '차변대변모두' 또는 공백이면 전체 유지

        if sub.empty:
            continue

        # 날짜순 정렬
        if COL_DATE in sub.columns:
            sub = sub.sort_values(COL_DATE)

        # 시트명: 상세거래_계정명_차변/대변/차변대변모두
        col_f_label = '차변' if col_f_norm in ('차변', '차변만') \
                      else '대변' if col_f_norm in ('대변', '대변만') \
                      else '차변대변모두'
        acct_short = re.sub(r'[^가-힣a-zA-Z0-9]', '', acct_name)[:12]
        base = _safe_sheet(f'상세거래_{acct_short}_{col_f_label}')
        sheet_name = base
        suffix = 2
        while sheet_name in all_results:
            sheet_name = f'{base[:28]}_{suffix}'
            suffix += 1

        all_results[sheet_name] = sub.reset_index(drop=True)

    if not all_results:
        return {'계정별상세내역': pd.DataFrame({'안내': ['해당 계정의 전표 내역이 없습니다.']})}

    return all_results


# =============================================================================
# 4. 분석 레지스트리  {번호: (이름, 함수)}
# =============================================================================
ANALYSIS_REGISTRY: dict = {
    2:  ('거래처비교',      analyze_client_comparison),
    3:  ('벤포드분석',      analyze_benford),
    4:  ('데이터개요',      analyze_data_overview),
    5:  ('계정명리스트',    analyze_account_list),
    6:  ('사원별집계',      analyze_employee_summary),
    7:  ('일자차이분석',    analyze_date_difference),
    8:  ('상대계정분석',    analyze_counterpart),
    9:  ('키워드검색',      analyze_keyword_search),
    10: ('라운드넘버',      analyze_round_numbers),
    11: ('특수관계자분석',  analyze_related_party),
    12: ('자산부채교차',    analyze_asset_liability_cross),
    13: ('매출비용교차',    analyze_revenue_expense_cross),
    14: ('심층분석',        analyze_top_accounts),
    15: ('AI계정별분석',    analyze_ai_preparation),
    16: ('헤더확인',        analyze_header_check),
    17: ('거래처분석',      analyze_client_detail),
    18: ('벤포드이탈',      analyze_benford_deviation),
    19: ('월별전계정분석',  analyze_monthly_full_account),
    20: ('잔액증감분석',    analyze_balance_movement),
    21: ('총계정원장',      analyze_general_ledger),
    22: ('은행조회서완전성', analyze_bank_confirmation),
    23: ('계정별상세내역',  analyze_account_transaction_detail),
}


# =============================================================================
# 5. Task List 읽기
# =============================================================================

def resolve_paths(company_name: str) -> dict:
    company_dir = os.path.join(BASE_DIR, company_name)
    return {
        'company_dir': company_dir,
        'task_list':   os.path.join(company_dir, f'task_list_{company_name}.xlsx'),
        'output':      os.path.join(company_dir, 'results'),
    }

def load_active_tasks(task_list_path: str) -> list:
    if not os.path.isfile(task_list_path):
        raise FileNotFoundError(f'task_list 없음: {task_list_path}')
    xl = pd.ExcelFile(task_list_path)
    df = None
    if TASK_MASTER_SHEET in xl.sheet_names:
        df = pd.read_excel(task_list_path, sheet_name=TASK_MASTER_SHEET)
    else:
        for sh in xl.sheet_names:
            tmp = pd.read_excel(task_list_path, sheet_name=sh)
            if all(any(k in str(c) for c in tmp.columns) for k in ('번호','명','여부')):
                df = tmp; break
    if df is None:
        raise ValueError(f"'분석목록' 시트를 찾을 수 없음: {task_list_path}")

    col_no     = next((c for c in df.columns if '번호' in str(c)), None)
    col_nm     = next((c for c in df.columns if str(c).strip().endswith('명')
                       and '번호' not in str(c) and '설명' not in str(c)), None)
    col_flag   = next((c for c in df.columns if '여부' in str(c)), None)
    col_period = next((c for c in df.columns if '대상' in str(c)), None)
    if not all([col_no, col_nm, col_flag]):
        raise ValueError(f'분석번호/분석명/실행여부 컬럼 없음. 실제 컬럼: {df.columns.tolist()}')

    flag   = df[col_flag].astype(str).str.strip().str.upper()
    active = df[flag.isin(['Y','O'])].dropna(subset=[col_no])
    tasks  = [
        (int(row[col_no]), str(row[col_nm]).strip(),
         str(row[col_period]).strip() if col_period and str(row[col_period]).strip() not in ('nan', '') else '당기')
        for _, row in active.iterrows()
    ]
    print(f'  [태스크] {len(tasks)}개: {[f"{n}_{nm}[{p}]" for n,nm,p in tasks]}')
    return tasks

def load_analysis_params(task_list_path: str, analysis_name: str) -> list:
    xl = pd.ExcelFile(task_list_path)
    for candidate in [analysis_name, f'{analysis_name}_파라미터']:
        if candidate not in xl.sheet_names: continue
        df = pd.read_excel(task_list_path, sheet_name=candidate).dropna(how='all')
        if '실행여부' in df.columns:
            flag = df['실행여부'].astype(str).str.strip().str.upper()
            df   = df[flag.isin(['Y','O'])].copy()
        params = df.to_dict('records')
        print(f'    └ 파라미터 시트 [{candidate}]: {len(params)}행')
        return params
    return [{}]   # 파라미터 시트 없음 → 함수 내 기본값 사용


# =============================================================================
# 6. 결과 저장
# =============================================================================

def load_settings(task_list_path: str) -> dict:
    """task_list settings 시트에서 회사정보(ClientName/StartDate/EndDate 등) 읽기."""
    try:
        xl = pd.ExcelFile(task_list_path)
        if 'settings' not in xl.sheet_names:
            return {}
        df = pd.read_excel(task_list_path, sheet_name='settings', header=None)
        settings = {}
        for _, row in df.iterrows():
            key = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ''
            val = str(row.iloc[1]).strip() if len(row) > 1 and pd.notna(row.iloc[1]) else ''
            if not key or not val or key in ('항목', 'key', 'Key'):
                continue
            # 날짜 형식 정리: '2026-01-01 00:00:00' → '2026-01-01'
            if 'Date' in key or '일자' in key:
                val = val.split(' ')[0].split('T')[0]
            settings[key] = val
        return settings
    except Exception:
        return {}


def save_results(results: dict, output_dir: str, company_name: str,
                 settings: dict = None) -> str:
    os.makedirs(output_dir, exist_ok=True)
    out_path = os.path.join(output_dir, f'분석결과_{company_name}.xlsx')

    settings      = settings or {}
    client_name   = settings.get('ClientName', company_name)
    start_date    = settings.get('StartDate', '')
    end_date      = settings.get('EndDate', '')
    header_line1  = f'회사명: {client_name}'
    header_line2  = f'분석기간: {start_date} ~ {end_date}' if start_date else ''

    # 특수 키 사전 추출 (ExcelWriter에 넘기지 않음)
    benford_images = results.pop('_benford_images', None)
    decoder        = results.pop('_암호해독표', None)

    # startrow=2: 1~2행을 비워두고 3행부터 컬럼헤더+데이터 기록
    with pd.ExcelWriter(out_path, engine='openpyxl') as writer:
        for sheet, df in results.items():
            sname = _safe_sheet(sheet)
            if df is None or (isinstance(df, pd.DataFrame) and df.empty):
                pd.DataFrame([['결과 없음']]).to_excel(
                    writer, sheet_name=sname, index=False, header=False, startrow=2)
            elif isinstance(df, pd.DataFrame):
                df.to_excel(writer, sheet_name=sname, index=False, startrow=2)

    # 회사 정보 상단 삽입 + 벤포드 차트 삽입 (openpyxl로 한 번에 처리)
    wb = openpyxl.load_workbook(out_path)
    from openpyxl.styles import Font
    bold = Font(bold=True)
    for ws in wb.worksheets:
        ws.cell(1, 1).value = header_line1
        ws.cell(1, 1).font  = bold
        if header_line2:
            ws.cell(2, 1).value = header_line2

    if benford_images:
        for _acct, _dir, img_buf in benford_images:
            if not img_buf: continue
            sname = _safe_sheet(f'벤포드_{_acct}_{_dir}')
            if sname in wb.sheetnames:
                wb[sname].add_image(XLImage(img_buf), 'K4')

    wb.save(out_path)

    # 암호해독표 → 별도 파일
    if decoder is not None:
        key_path = os.path.join(output_dir, f'암호해독표_{company_name}.xlsx')
        decoder.to_excel(key_path, index=False)
        print(f'  [암호해독표] {key_path}')

    return out_path


# =============================================================================
# 7. 메인
# =============================================================================

def main():
    parser = argparse.ArgumentParser(description='분개장 분석 자동화')
    parser.add_argument('company', nargs='?', help='고객사 이름 (예: sejoong)')
    parser.add_argument('--task', type=int, nargs='+', metavar='N', help='실행할 분석 번호 (예: --task 21  또는  --task 3 8 21)')
    args = parser.parse_args()
    company_name = (args.company or input('고객사 이름: ')).strip()
    if not company_name:
        print('[오류] 고객사 이름이 비어 있습니다.'); sys.exit(1)

    global _COMPANY_DIR
    print(f'\n{"="*60}\n  분개장 분석 자동화 — {company_name}\n{"="*60}')
    paths = resolve_paths(company_name)
    _COMPANY_DIR = paths['company_dir']

    # 1) 태스크 리스트
    try:
        active_tasks = load_active_tasks(paths['task_list'])
    except (FileNotFoundError, ValueError) as e:
        print(f'[오류] {e}'); sys.exit(1)
    if args.task:
        active_tasks = [(n, nm, p) for n, nm, p in active_tasks if n in args.task]
        print(f'  [필터] --task {args.task} → {len(active_tasks)}개 실행')
    if not active_tasks:
        print('실행할 분석이 없습니다 (Y/O 항목 없음).'); sys.exit(0)

    # 2) 분개장 로드
    print('\n[분개장 로드]')
    df = load_data(paths['company_dir'])
    if df is None:
        print('[오류] data/current 또는 data/previous 폴더에 분개장 파일이 없습니다.')
        sys.exit(1)
    df = _apply_company_preprocess(df, company_name)   # 회사별 전용 전처리
    df = _preprocess_df(df)                            # 공통 표준 전처리
    gc = _get_gubun_col(df)
    print(f'  총 {len(df):,}행'
          + (f' (당기: {(df[gc]=="당기").sum():,}건 / 전기: {(df[gc]=="전기").sum():,}건)' if gc else '')
          + f'\n  컬럼: {df.columns.tolist()}')

    # 3) 분석 순차 실행
    print('\n[분석 실행]')
    all_results: dict = {}
    for task_no, task_name, 분석대상 in active_tasks:
        if task_no not in ANALYSIS_REGISTRY:
            print(f'  [{task_no:>3}] {task_name:<22} → 등록된 함수 없음 (건너뜀)')
            continue
        _, func = ANALYSIS_REGISTRY[task_no]
        params_list = load_analysis_params(paths['task_list'], task_name)
        # 분석대상(당기/전기/전체)에 따라 df 슬라이싱
        if '구분' in df.columns and 분석대상 in ('당기', '전기'):
            task_df = df[df['구분'] == 분석대상].copy()
        else:
            task_df = df
        print(f'  [{task_no:>3}] {task_name} [{분석대상} {len(task_df):,}행]', flush=True)
        try:
            result = func(task_df, params_list)
            if isinstance(result, dict):
                for sname, sub_df in result.items():
                    all_results[sname] = sub_df
            elif isinstance(result, pd.DataFrame):
                all_results[_safe_sheet(task_name)] = result
            print(f'       → 시트 {len(result) if isinstance(result, dict) else 1}개 생성')
        except Exception as e:
            import traceback
            print(f'       ⚠️ 오류: {e}')
            traceback.print_exc()

    # 4) 결과 저장
    print('\n[결과 저장]')
    settings = load_settings(paths['task_list'])
    out_path = save_results(all_results, paths['output'], company_name, settings)
    print(f'\n  ✅ 완료: {out_path}')
    print(f'  시트 수: {len(all_results)}개')


if __name__ == '__main__':
    main()
