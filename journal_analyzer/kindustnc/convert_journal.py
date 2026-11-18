#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
kindustnc 전용 분개장 통합 (본점 분개장 + 지점 분개장 → 통합 분개장 1개)

원본 (더존 위하고, 2026-11-18 확인):
  - data/{current|previous}/ 에 '..분개장_본점_..xlsx', '..분개장_지점_..xlsx' 2개
  - 본점: 10열(전표일자/전표번호/구분/계정과목코드/계정과목/차변/대변/거래처/거래처코드/적요)
  - 당기 지점: 같은 10열이나 머리글이 'Code'/'Code.1'이고 K~P열에 보조 수식 열이 더 있음
    → 앞 10열만 쓰고 머리글을 본점 기준으로 맞춘다.

통합 시 처리:
  1) 지점 합산전표 제외 — 당기 본점 분개장에는 지점 9월말 잔액을 한 번에 넣은 합산전표
     (적요 '합계잔액시산표9월말잔액', 거래처 '지점')가 있다. 통합 계정별원장은 이 전표를
     포함한 본점 장부다. 지점 거래를 따로 합치므로 합산전표를 그대로 두면 지점 금액이
     두 번 잡힌다 → 제외한다. 제외 전에 합산전표 손익계정 금액 = 지점 분개장 손익 순액인지
     대사해 출력한다(지점 분개장이 빠짐없는지 확인). 재무상태표 계정 차이는 지점 기초잔액.
  2) 전표번호 충돌 방지 — 본점·지점 모두 '일자+순번'(2026010200001)으로 번호를 매겨 당기
     297건 중 288건이 본점 번호와 겹친다. 엔진 전표그룹키는 일자+전표번호라 그대로 두면
     다른 전표가 하나로 묶인다 → 지점 전표번호 앞에 '지'를 붙인다.
  3) '사업장' 열(본점/지점) 추가.
  4) 본지점(980) 계정은 통합하면 상계되어야 한다 → 통합 후 잔액(순액)을 출력한다.

원본은 data/raw/{period}/ 로 옮기고 '{당기|전기}_kindustnc분개장_통합_{yy}년.xlsx'만 남긴다
(엔진 load_data는 data/{period}의 xlsx를 모두 읽으므로 원본이 남아 있으면 중복 로드됨).
원본이 바뀌지 않았으면 다시 만들지 않는다.

사용법: main_analyzer.py kindustnc 실행 시 분개장 로드 전에 자동 실행됨(_run_company_converter).
        단독 실행도 가능: python journal_analyzer/kindustnc/convert_journal.py
"""
import os
import re
import shutil
import sys

import pandas as pd

_DIR = os.path.dirname(os.path.abspath(__file__))
_PERIODS = ('current', 'previous')
_COLS = ['전표일자', '전표번호', '구분', '계정과목코드', '계정과목', '차변', '대변', '거래처', '거래처코드', '적요']
_BRANCH_PREFIX = '지'
_INTER_OFFICE_CODE = '980'          # 본지점
_MERGE_MEMO = '합계잔액시산표'       # 본점 분개장의 지점 합산전표 적요
_MERGE_PARTNER = '지점'
_SRC_RE = re.compile(r'분개장_(본점|지점)_(\d{2})년\.xlsx$')


def _num(s: pd.Series) -> pd.Series:
    return pd.to_numeric(s.astype(str).str.replace(',', '', regex=False).str.strip(),
                         errors='coerce').fillna(0)


def _read(path: str) -> pd.DataFrame:
    df = pd.read_excel(path, dtype=str, header=0)
    df = df.iloc[:, :len(_COLS)].copy()
    df.columns = _COLS
    df = df[df['전표일자'].notna() & df['전표번호'].notna()]
    df = df[~df['전표일자'].astype(str).str.strip().str.startswith('합계')]   # 위하고 합계행
    df = df[df['전표번호'].astype(str).str.strip() != '']
    df['전표일자'] = df['전표일자'].astype(str).str.strip().str[:10]
    for c in ('전표번호', '계정과목코드', '거래처코드'):
        df[c] = df[c].where(df[c].isna(), df[c].astype(str).str.strip())
    df['차변'], df['대변'] = _num(df['차변']), _num(df['대변'])
    return df.reset_index(drop=True)


def _check_balanced(df: pd.DataFrame, label: str) -> bool:
    g = df.groupby('전표번호')[['차변', '대변']].sum()
    bad = g[(g['차변'] - g['대변']).abs() > 0.5]
    if len(bad):
        print(f'  [오류] {label} 차대 불일치 전표 {len(bad)}건: {list(bad.index[:5])}')
        return False
    return True


_BS_9XX = {'962', '964', _INTER_OFFICE_CODE}   # 위하고 9xx 중 재무상태표 계정(보증금·본지점)


def _is_pl(code) -> bool:
    if str(code) in _BS_9XX:
        return False
    try:
        return int(code) >= 400
    except (TypeError, ValueError):
        return False


def merge(head_path: str, branch_path: str, label: str) -> pd.DataFrame:
    head, branch = _read(head_path), _read(branch_path)
    print(f'  [{label}] 본점 {len(head):,}행 차변 {head["차변"].sum():,.0f} / '
          f'지점 {len(branch):,}행 차변 {branch["차변"].sum():,.0f}')
    if not (_check_balanced(head, f'{label} 본점') and _check_balanced(branch, f'{label} 지점')):
        raise ValueError('차대 불일치 전표 있음')

    # 1) 지점 합산전표 제외
    memo = head['적요'].fillna('').astype(str)
    partner = head['거래처'].fillna('').astype(str).str.strip()
    merge_nos = set(head.loc[memo.str.contains(_MERGE_MEMO) & (partner == _MERGE_PARTNER), '전표번호'])
    if merge_nos:
        m = head[head['전표번호'].isin(merge_nos)]
        print(f'  [{label}] 지점 합산전표 제외: {sorted(merge_nos)} ({len(m)}행, 차변 {m["차변"].sum():,.0f})')
        # 손익계정: 합산전표 금액 = 지점 분개장 순액이어야 함 (지점 분개장 누락 검증)
        mnet = m.assign(net=m['차변'] - m['대변']).groupby('계정과목코드')['net'].sum()
        bnet = branch.assign(net=branch['차변'] - branch['대변']).groupby('계정과목코드')['net'].sum()
        codes = sorted(set(mnet.index) | set(bnet.index))
        diff = pd.DataFrame({'합산전표': mnet.reindex(codes).fillna(0),
                             '지점분개장': bnet.reindex(codes).fillna(0)})
        diff['차이'] = diff['합산전표'] - diff['지점분개장']
        pl = diff[[_is_pl(c) for c in diff.index]]
        pl_bad = pl[pl['차이'].abs() > 0.5]
        print(f'  [{label}] 손익계정 합산전표 vs 지점 분개장: {len(pl)}개 계정 중 차이 {len(pl_bad)}개'
              + ('' if pl_bad.empty else f'\n{pl_bad.to_string()}'))
        bs = diff[[not _is_pl(c) and c != _INTER_OFFICE_CODE for c in diff.index]]
        bs = bs[bs['차이'].abs() > 0.5]
        if not bs.empty:
            print(f'  [{label}] 재무상태표 계정 차이(= 지점 기초잔액, 정상):\n{bs.to_string()}')
        head = head[~head['전표번호'].isin(merge_nos)]

    # 2) 전표번호 구분 + 3) 사업장
    branch = branch.assign(전표번호=_BRANCH_PREFIX + branch['전표번호'].astype(str))
    out = pd.concat([head.assign(사업장='본점'), branch.assign(사업장='지점')], ignore_index=True)
    out = out.sort_values(['전표일자', '사업장', '전표번호'], kind='stable').reset_index(drop=True)

    # 검증: 통합 합계 = 본점(합산전표 제외) + 지점
    exp_d, exp_c = head['차변'].sum() + branch['차변'].sum(), head['대변'].sum() + branch['대변'].sum()
    if abs(out['차변'].sum() - exp_d) > 0.5 or abs(out['대변'].sum() - exp_c) > 0.5:
        raise ValueError('통합 합계 불일치')
    if not _check_balanced(out, f'{label} 통합'):
        raise ValueError('통합 차대 불일치')
    if out['전표번호'].duplicated().any() and \
            (out.groupby('전표번호')['사업장'].nunique() > 1).any():
        raise ValueError('본점·지점 전표번호 충돌 남음')

    # 4) 본지점 상계 확인
    io = out[out['계정과목코드'] == _INTER_OFFICE_CODE]
    io_net = io['차변'].sum() - io['대변'].sum()
    print(f'  [{label}] 통합 {len(out):,}행 차변 {out["차변"].sum():,.0f} = 대변 {out["대변"].sum():,.0f}, '
          f'본지점(980) 순액 {io_net:,.0f}' + (' (상계 완료)' if abs(io_net) < 0.5 else ' ← 미상계 잔액 확인 필요'))
    if abs(io_net) >= 0.5:
        by_site = io.assign(net=io['차변'] - io['대변']).groupby('사업장')['net'].sum()
        print(f'           사업장별 본지점 순액: {by_site.to_dict()}')
    return out[_COLS + ['사업장']]


def main() -> int:
    for period in _PERIODS:
        data_dir = os.path.join(_DIR, 'data', period)
        raw_dir = os.path.join(_DIR, 'data', 'raw', period)
        if not os.path.isdir(data_dir):
            continue
        # 1) data/{period}의 본점·지점 원본 → data/raw/{period}
        for f in sorted(os.listdir(data_dir)):
            if f.startswith('~$') or not _SRC_RE.search(f):
                continue
            os.makedirs(raw_dir, exist_ok=True)
            dest = os.path.join(raw_dir, f)
            if os.path.exists(dest):
                os.remove(dest)      # 새로 올린 원본으로 교체
            shutil.move(os.path.join(data_dir, f), dest)
            print(f'[{period}] 원본 이동: {f} → data/raw/{period}/')
        if not os.path.isdir(raw_dir):
            continue
        # 2) 본점+지점 → 통합
        pairs = {}
        for f in sorted(os.listdir(raw_dir)):
            m = _SRC_RE.search(f)
            if m and not f.startswith('~$'):
                pairs.setdefault(m.group(2), {})[m.group(1)] = os.path.join(raw_dir, f)
        for yy, p in pairs.items():
            if set(p) != {'본점', '지점'}:
                print(f'[오류] {period} {yy}년: 본점·지점 분개장이 모두 있어야 함 (있는 것: {sorted(p)})')
                return 1
            prefix = '당기' if period == 'current' else '전기'
            out_path = os.path.join(data_dir, f'{prefix}_kindustnc분개장_통합_{yy}년.xlsx')
            if os.path.isfile(out_path) and \
                    os.path.getmtime(out_path) >= max(os.path.getmtime(x) for x in p.values()):
                continue
            print(f'[{period}] 본점+지점 분개장 통합 ({yy}년)')
            try:
                out = merge(p['본점'], p['지점'], prefix)
            except ValueError as e:
                print(f'[오류] {period} 통합 실패: {e}')
                return 1
            out.to_excel(out_path, index=False)
            print(f'[{period}] 저장: {os.path.basename(out_path)}')
    return 0


if __name__ == '__main__':
    sys.exit(main())
