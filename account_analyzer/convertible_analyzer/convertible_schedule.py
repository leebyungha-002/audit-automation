# -*- coding: utf-8 -*-
"""전환사채(CB) 전환권평가·회계처리 검증앱.

input_data/convertible_<company>_information_fy<year>.xlsx 를 읽어
  ① 발행시 부채·자본(전환권대가) 분리, 유효이자율 상각표 (K-IFRS 1032, 전환권 자본분류)
  ② 회계연도 기초·기말 장부금액, 이자비용, 전환·조기상환 처리 → 회사계상 대사
  ③ 평가시점별 공정가치 독립 재계산(TF 이항격자: 전환권·조기상환청구권·매도청구권 분리)
     → 외부평가보고서 금액 대사, 민감도
  ④ 검토사항(분류·유동성·매도청구권 보유자·검증합계)
를 output/convertible_schedule_<company>_<fy>.xlsx 로 출력한다.

실행: python convertible_schedule.py graphy --fiscal-year 2026 [--interim-month 9]
"""
import os
import re
import sys
import glob
import argparse
import calendar
from datetime import date, datetime, timedelta

import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.path.insert(0, os.path.join(HERE, '..', '..', 'shared_modules'))
import convertible_engine as E  # noqa: E402

try:
    import run_stamp
except ImportError:
    run_stamp = None

if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8', line_buffering=True)

INPUT_DIR = os.path.join(HERE, 'input_data')
OUTPUT_DIR = os.path.join(HERE, 'output')
TOL_RATE = 0.01      # 대사 차이 유의 기준(1%)
TOL_AMT = 1_000      # 원단위 반올림 허용


# ─── 입력 ────────────────────────────────────────────────────────────────────

def _num(v, default=None):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return default
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r'[,\s원%]', '', str(v))
    try:
        return float(s) if s else default
    except ValueError:
        return default


def _pct(v, default=None):
    x = _num(v, None)
    return default if x is None else x / 100.0


def _date(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    try:
        return pd.Timestamp(str(v)).date()
    except Exception:
        return None


def _read_sheet(path, name):
    """1행 헤더, 2행 설명(건너뜀), 3행부터 데이터. '예시)' 회차·빈 회차 제외."""
    try:
        df = pd.read_excel(path, sheet_name=name, header=0)
    except ValueError:
        return []
    df = df.iloc[1:]
    rows = []
    for rec in df.to_dict('records'):
        k = str(rec.get('회차') or '').strip()
        if not k or k == 'nan' or k.startswith('예시'):
            continue
        rec['회차'] = k
        rows.append(rec)
    return rows


def find_input(company=None, file=None):
    if file:
        p = file if os.path.isabs(file) else os.path.join(INPUT_DIR, file)
        if not os.path.exists(p):
            sys.exit(f'[오류] 입력 파일 없음: {p}')
        return p
    pat = f'convertible_{company}_information_fy*.xlsx' if company else 'convertible_*_information_fy*.xlsx'
    cands = [p for p in glob.glob(os.path.join(INPUT_DIR, pat)) if not os.path.basename(p).startswith('~$')]
    if not cands:
        sys.exit(f'[오류] 입력 파일 없음: input_data/{pat}\n'
                 f'  → python build_template.py 로 템플릿을 만든 뒤 복사해서 작성하세요.')
    cands.sort(key=os.path.getmtime, reverse=True)
    if len(cands) > 1:
        print(f'[주의] 입력 파일 {len(cands)}개 — 최근 파일 사용: {os.path.basename(cands[0])}')
    return cands[0]


def build_terms(path):
    """회차 → terms dict."""
    issues = _read_sheet(path, '발행조건')
    puts = _read_sheet(path, '조기상환Put')
    calls = _read_sheet(path, '매도청구Call')
    out = {}
    for r in issues:
        k = r['회차']
        issue, mat = _date(r.get('발행일')), _date(r.get('만기일'))
        if not issue or not mat:
            print(f'[경고] {k}: 발행일/만기일 누락 — 건너뜀')
            continue
        cpn = _pct(r.get('표면이자율(%)'), 0.0)
        freq = int(_num(r.get('이자지급주기(개월)'), 0) or 0)
        ytm = _pct(r.get('만기보장수익률(%)'), 0.0)
        red = _pct(r.get('만기상환율(%)'))
        red_src = '입력'
        if red is None:
            red = E.redemption_ratio_from_yield(issue, mat, cpn, freq, ytm)
            red_src = f'YTM {ytm:.4%}로 계산'
        flows = E.contract_cashflows(issue, mat, cpn, freq, red)
        conv_price = _num(r.get('전환가격(원)'))
        t = {
            '회차': k, '발행일': issue, '만기일': mat,
            '액면': _num(r.get('액면총액(원)'), 0.0), '발행가액': _num(r.get('발행가액(원)'), 0.0),
            '거래비용': _num(r.get('거래비용(원)'), 0.0), '표면이자율': cpn, '이자주기': freq,
            'YTM': ytm, '만기상환율': red, '만기상환율_출처': red_src,
            '전환가격': conv_price, '전환비율': 1.0 / conv_price if conv_price else 0.0,
            '주당액면': _num(r.get('1주당액면(원)'), 0.0),
            '전환청구시작일': _date(r.get('전환청구시작일')) or issue,
            '전환청구종료일': _date(r.get('전환청구종료일')) or mat,
            '시장이자율': _pct(r.get('일반사채시장이자율(%)')),
            '부채요소금액': _num(r.get('부채요소금액(원)')),
            '리픽싱': str(r.get('리픽싱(Y/N)') or '').strip().upper() == 'Y',
            '현금흐름': flows,
            '현금흐름_이자': [(d, a) for d, a in flows],  # 격자용(만기분 포함 전체)
            'puts': [], 'calls': [], 'put_rows': [], 'call_rows': [],
        }
        # 격자 만기 처리: 만기상환액은 redemption 인자로, 만기 표면이자만 쿠폰으로
        last_cpn = cpn * freq / 12.0 if freq else 0.0
        t['현금흐름_이자'] = [(d, a) for d, a in flows if d != mat] + ([(mat, last_cpn)] if last_cpn else [])
        out[k] = t

    for r in puts:
        t = out.get(r['회차'])
        d = _date(r.get('행사일'))
        if not t or not d:
            continue
        pr = _pct(r.get('상환율(%)'))
        y = _pct(r.get('조기상환수익률(%)'))
        if pr is None and y is not None:
            pr = E.redemption_ratio_from_yield(t['발행일'], d, t['표면이자율'], t['이자주기'], y)
        if pr is None:
            print(f"[경고] {t['회차']} Put {d}: 상환율·수익률 모두 없음 — 건너뜀")
            continue
        t['puts'].append((d, pr))
        t['put_rows'].append({'행사일': d, '조기상환수익률': y, '상환율': pr})

    for r in calls:
        t = out.get(r['회차'])
        d0, d1 = _date(r.get('행사시작일')), _date(r.get('행사종료일'))
        if not t or not d0 or not d1:
            continue
        fixed = _pct(r.get('행사가율(%)'))
        y = _pct(r.get('행사가수익률(%)'))
        if fixed is not None:
            fn = (lambda f: (lambda dd: f))(fixed)
        elif y is not None:
            fn = (lambda tt, yy: (lambda dd: E.redemption_ratio_from_yield(
                tt['발행일'], dd, tt['표면이자율'], tt['이자주기'], yy)))(t, y)
        else:
            print(f"[경고] {t['회차']} Call: 행사가율·수익률 모두 없음 — 건너뜀")
            continue
        f = _pct(r.get('행사가능비율(%)'), 0.0)
        holder = str(r.get('보유자') or '').strip()
        method = str(r.get('평가방법') or '').strip()
        method = '콜행사반영' if method == '콜행사반영' else '별도격자'
        asset = str(r.get('자산인식(Y/N)') or '').strip().upper() == 'Y'
        disc = str(r.get('할인방식') or '').strip()
        disc = disc if disc in E.CALL_DISCOUNTS else 'TF분할'
        pay = _date(r.get('지급기일'))
        row = {'행사시작일': d0, '행사종료일': d1, '지급기일': pay, '행사가수익률': y, '행사가율': fixed,
               '행사가능비율': f, '보유자': holder, '평가방법': method, '할인방식': disc,
               '자산인식': 'Y' if asset else 'N'}
        t['call_rows'].append(row)
        if pay:   # 차수별 청구기간·지급기일이 정해진 경우 → 한 회차의 행들을 묶어 버뮤다형 1건으로
            row['_가격'] = fn(pay)
            continue
        t['calls'].append({'시작일': d0, '종료일': d1, '가격함수': fn, '행사가능비율': f, '보유자': holder,
                           '평가방법': method, '할인방식': disc, '자산인식': asset})
    for t in out.values():
        brs = [r for r in t['call_rows'] if r.get('지급기일')]
        if brs:
            b0 = brs[0]
            sched = sorted((r['지급기일'], r['_가격']) for r in brs)
            t['calls'].append({'시작일': min(r['행사시작일'] for r in brs), '종료일': sched[-1][0],
                               '가격함수': (lambda sc: (lambda dd: next((pr for d, pr in sc if d >= dd), sc[-1][1])))(sched),
                               '행사일정': sched, '행사가능비율': b0['행사가능비율'], '보유자': b0['보유자'],
                               '평가방법': '별도격자', '할인방식': b0['할인방식'], '자산인식': b0['자산인식'] == 'Y'})
    for t in out.values():
        t['puts'].sort()
    return out


TENOR_COLS = [('3M', 0.25), ('6M', 0.5), ('9M', 0.75), ('1Y', 1), ('1.5Y', 1.5), ('2Y', 2), ('2.5Y', 2.5),
              ('3Y', 3), ('4Y', 4), ('5Y', 5), ('7Y', 7), ('10Y', 10)]


def load_curves(path):
    """'금리기간구조' 시트 → {(회차, 평가기준일): {'무위험': Curve, '위험': Curve}}."""
    out = {}
    for r in _read_sheet(path, '금리기간구조'):
        d = _date(r.get('평가기준일'))
        kind = str(r.get('구분') or '').strip()
        kind = '무위험' if '무위험' in kind or '국고' in kind else ('위험' if kind else '')
        if not d or not kind:
            continue
        pts = [(tn, _pct(r.get(col))) for col, tn in TENOR_COLS if _pct(r.get(col)) is not None]
        if len(pts) < 2:
            continue
        out.setdefault((r['회차'], d), {})[kind] = E.Curve([p[0] for p in pts], [p[1] for p in pts])
    return out


def group_by(rows, key='회차'):
    g = {}
    for r in rows:
        g.setdefault(r[key], []).append(r)
    return g


def fy_bounds(fy: int, fiscal_month: int, interim_month=None):
    end_m = interim_month or fiscal_month
    end_y = fy if end_m <= fiscal_month else fy - 1
    end = date(end_y, end_m, calendar.monthrange(end_y, end_m)[1])
    prior = date(fy - 1, fiscal_month, calendar.monthrange(fy - 1, fiscal_month)[1])
    return prior + timedelta(days=1), end, prior


# ─── 회차별 계산 ─────────────────────────────────────────────────────────────

def _shift(r, dx):
    """금리(단일 또는 Curve)를 dx 만큼 평행이동."""
    if isinstance(r, E.Curve):
        if r.flat is not None:
            return E.Curve(flat=r.flat + dx)
        c = E.Curve.__new__(E.Curve)
        c.__dict__.update(r.__dict__)
        c.logdf = r.logdf - dx * r.t   # 연속복리 근사 평행이동
        return c
    return r + dx


def _quarter_ends(d0: date, d1: date, fiscal_month: int) -> list:
    """(d0, d1) 사이 분기말 [(일자, 회계연도, 구분)]. 구분 = 1분기/반기/3분기/결산."""
    out, y, m = [], d0.year, d0.month
    while True:
        e = date(y, m, calendar.monthrange(y, m)[1])
        if e >= d1:
            break
        q = (m - fiscal_month) % 12
        if e > d0 and q % 3 == 0:
            fy = y if m <= fiscal_month else y + 1
            out.append((e, fy, {3: '1분기', 6: '반기', 9: '3분기', 0: '결산'}[q]))
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    return out


def adj_schedule(t, state, flows, eir, ev_res, comp, fiscal_month, fy_start, fy_end):
    """전환권조정 분기별 상각표 (발행일 ~ 만기, 2026-11-11 blue sky 요청).

    전환권조정 = 미상환액면 + 상환할증금 − 장부금액. 상각액 = 유효이자 − 표면이자.
    전환·상환 감소 = 기초 − 상각 − 기말 (이벤트로 액면·할증금과 함께 제거된 조정액).
    원 단위: 잔액을 반올림하고 상각액을 끝수 조정 값으로 맞춰 만기 잔액이 정확히 0이 되게 한다.
    """
    def fy_of(d):
        return d.year if d.month <= fiscal_month else d.year + 1

    pts = [(t['발행일'], fy_of(t['발행일']), '발행')] + _quarter_ends(t['발행일'], t['만기일'], fiscal_month)
    pts.append((t['만기일'], fy_of(t['만기일']), '만기'))
    rows = []
    for (a, _, _), (b, fyb, tag) in zip(pts[:-1], pts[1:]):
        if state(a)['미상환액면'] <= 0:
            break
        ca = comp(a)
        cb = ({kk: 0.0 for kk in ca} if b == t['만기일'] else comp(b))
        intr = E.interest_between(state, flows, eir, a, b, ev_res)
        paid = E.coupons_paid_between(state, flows, a, b, t['만기상환율'], t['만기일'])
        evs = [e for e in ev_res if a < e['일자'] <= b]
        adj_a, adj_b = round(ca['전환권조정']), round(cb['전환권조정'])
        dec = round(ca['전환권조정'] - (intr - paid) - cb['전환권조정'])
        amort = adj_a - dec - adj_b                        # 끝수 조정 포함 상각액
        rows.append({'회계연도': fyb, '구분': tag, '기간시작': a, '기간종료': b,
                     '기초 전환권조정': adj_a, '상각액': amort, '전환·상환 감소': dec, '기말 전환권조정': adj_b,
                     '유효이자': paid + amort, '표면이자': paid,
                     '사채액면': cb['미상환액면'], '상환할증금': cb['상환할증금'], '장부금액': cb['장부금액'],
                     '비고': ', '.join(f"{e['일자']} {e['구분']}" for e in evs),
                     '당기': fy_start <= b <= fy_end})
    return rows


def process(t, events_raw, vals_raw, books_raw, fy_start, fy_end, prior_end, steps, curves=None):
    k = t['회차']
    res = {'회차': k, 'terms': t, 'checks': []}
    chk = res['checks']
    flows = t['현금흐름']
    curves = curves or {}
    vals, sens, issue_val = [], [], None

    def outstanding(dd):
        o = t['액면']
        for e in events_raw:
            ed = _date(e.get('일자'))
            if ed and ed <= dd and dd > t['발행일']:
                o -= _num(e.get('액면금액(원)'), 0.0)
        return max(o, 0.0)

    # ① 공정가치 재계산 (부채요소 측정에 발행일 평가액을 쓰므로 회계처리보다 먼저)
    for v in vals_raw:
        vd = _date(v.get('평가기준일'))
        if not vd or vd >= t['만기일']:
            continue
        cv = curves.get((k, vd), {})
        mkt = {'주가': _num(v.get('주가(원)')), '변동성': _pct(v.get('변동성(%)')),
               '무위험이자율': cv.get('무위험', _pct(v.get('무위험이자율(%)'), 0.0)),
               '신용위험조정할인율': cv.get('위험', _pct(v.get('신용위험조정할인율(%)'), 0.0)),
               '배당수익률': _pct(v.get('배당수익률(%)'), 0.0)}
        rate_src = '기간구조' if cv else '단일금리'
        if not mkt['주가'] or not mkt['변동성']:
            chk.append(('확인', f'평가시점 {vd}: 주가·변동성 누락 — 공정가치 재계산 생략'))
            continue
        o = outstanding(vd)
        c = E.value_components(t, mkt, vd, steps)
        row = {'평가기준일': vd, '미상환액면': o,
               **{kk: ({a: b * o for a, b in vv.items()} if isinstance(vv, dict) else vv * o) for kk, vv in c.items()},
               '시장변수': mkt,
               '금리': rate_src}
        ext = {}
        for col, key in (('외부평가_전환사채총가치(원)', '총가치'), ('외부평가_주계약(원)', '일반사채'),
                         ('외부평가_일반사채(원)', '일반사채'),
                         ('외부평가_전환권(원)', '전환권'), ('외부평가_조기상환청구권(원)', '조기상환청구권'),
                         ('외부평가_매도청구권(원)', '매도청구권')):
            ev = _num(v.get(col))
            if ev is not None:
                ext[key] = abs(ev) if key == '매도청구권' else ev  # 보고서는 콜을 (−)로 표시하기도 함
        row['외부'] = ext
        vals.append(row)
        for key, ev in ext.items():
            ours = row[key]
            if abs(ev - ours) > max(TOL_AMT, abs(ev) * 0.05):
                chk.append(('확인', f'{vd} {key}: 외부평가 {ev:,.0f} vs 재계산 {ours:,.0f} '
                                   f'(차이 {(ours - ev) / ev if ev else 0:+.1%}) — 가정·모형 차이 검토'))
        if vd == t['발행일'] and t['발행가액']:
            issue_val = row
            gap = row['총가치'] - t['발행가액']
            if abs(gap) > t['발행가액'] * 0.05:
                chk.append(('확인', f'발행일 재계산 공정가치가 발행가액과 {gap / t["발행가액"]:+.1%} 차이 — '
                                   '잔여법으로 자본요소에 흡수됨(K-IFRS 1032.31). 시장변수 적정성 검토'))
            if '총가치' in ext and abs(ext['총가치'] - t['발행가액']) > t['발행가액'] * 0.05:
                chk.append(('확인', f"평가보고서 CB 전체 공정가치 {ext['총가치']:,.0f}가 발행가액과 "
                                   f"{(ext['총가치'] - t['발행가액']) / t['발행가액']:+.1%} 차이 — 발행조건 적정성·Day-1 차이 검토"))
        # 민감도
        base = row['총가치']
        for label, mk in (('변동성 -5%p', dict(mkt, 변동성=max(mkt['변동성'] - 0.05, 0.01))),
                          ('변동성 +5%p', dict(mkt, 변동성=mkt['변동성'] + 0.05)),
                          ('신용할인율 -1%p', dict(mkt, 신용위험조정할인율=_shift(mkt['신용위험조정할인율'], -0.01))),
                          ('신용할인율 +1%p', dict(mkt, 신용위험조정할인율=_shift(mkt['신용위험조정할인율'], 0.01))),
                          ('주가 -10%', dict(mkt, 주가=mkt['주가'] * 0.9)),
                          ('주가 +10%', dict(mkt, 주가=mkt['주가'] * 1.1))):
            cc = E.value_components(t, mk, vd, steps)
            sens.append({'평가기준일': vd, '시나리오': label, '총가치': cc['총가치'] * o,
                         '전환권': cc['전환권'] * o, '총가치 변동': cc['총가치'] * o - base,
                         '변동률': (cc['총가치'] * o - base) / base if base else 0.0})
    res['valuation'] = vals
    res['sensitivity'] = sens

    # ② 발행시 분리 — 부채요소: 입력값 > 발행일 평가(주계약+조기상환권) > 시장이자율 할인
    if t['부채요소금액']:
        liab, liab_src = t['부채요소금액'], '입력값(부채요소금액)'
    elif issue_val:
        liab, liab_src = issue_val['일반사채_Put포함'], '발행일 재계산(주계약+조기상환권)'
    elif t['시장이자율'] is not None:
        liab = t['액면'] * E.pv(flows, t['발행일'], t['시장이자율'])
        liab_src = f"시장이자율 {t['시장이자율']:.3%} 할인(조기상환권 미반영)"
    else:
        chk.append(('오류', '부채요소금액·발행일 평가시점·일반사채시장이자율 모두 없음 — 부채요소 측정 불가'))
        return res
    call_asset = 0.0
    if any(c['자산인식'] for c in t['calls']):
        if issue_val:
            call_asset = issue_val['매도청구권']
        else:
            chk.append(('확인', '매도청구권 자산인식(Y)이나 발행일 평가시점 없음 — 콜 자산 0으로 처리'))
    split = E.initial_split(t['발행가액'], t['액면'], flows, t['발행일'], liab, t['거래비용'], call_asset)
    split['부채요소_출처'] = liab_src
    res['split'] = split
    # 평가보고서 기준 분리(대사용): 부채 = 주계약 + 조기상환권, 콜 자산 = 매도청구권
    if issue_val and {'일반사채', '조기상환청구권'} <= set(issue_val['외부']):
        ex = issue_val['외부']
        rl = ex['일반사채'] + ex['조기상환청구권']
        rc = ex.get('매도청구권', 0.0) if any(c['자산인식'] for c in t['calls']) else 0.0
        res['split_report'] = {'부채요소': rl, '매도청구권_자산': rc,
                               '자본요소': t['발행가액'] - rl + rc}
    if split['전환권대가'] < 0:
        chk.append(('확인', f"전환권대가가 음수({split['전환권대가']:,.0f}) — 시장이자율이 YTM보다 낮음, 입력 확인"))

    # ② 원장
    events = []
    for e in events_raw:
        d = _date(e.get('일자'))
        x = _num(e.get('액면금액(원)'), 0.0)
        if d and x:
            events.append({'일자': d, '구분': str(e.get('구분') or '').strip() or '전환',
                           '액면금액': x, '지급액': _num(e.get('지급액(원)'), 0.0)})
    state, ev_res = E.run_ledger(t['액면'], flows, t['발행일'], split['유효이자율'], split['전환권대가'],
                                 events, t['전환가격'], t['주당액면'], until=t['만기일'])
    res['events'] = ev_res
    eir = split['유효이자율']

    def comp(dd):
        st = state(dd)
        o = st['미상환액면']
        prem = o * max(t['만기상환율'] - 1.0, 0.0)
        return {'미상환액면': o, '상환할증금': prem, '전환권조정': o + prem - st['장부금액'],
                '장부금액': st['장부금액'], '전환권대가': st['전환권대가']}

    # 상각표: 발행일, 이자지급일, 결산일(전기말·당기말), 이벤트일
    pts = {t['발행일'], t['만기일']} | {d for d, _ in flows}
    pts |= {e['일자'] for e in ev_res}
    for d in (prior_end, fy_end):
        if t['발행일'] < d < t['만기일']:
            pts.add(d)
    pts = sorted(pts)
    sched = []
    for a, b in zip(pts[:-1], pts[1:]):
        if state(a)['미상환액면'] <= 0:
            break
        ca = comp(a)
        intr = E.interest_between(state, flows, eir, a, b, ev_res)
        paid = E.coupons_paid_between(state, flows, a, b, t['만기상환율'], t['만기일'])
        evs = [e for e in ev_res if e['일자'] == b]
        off = sum(e['부채제거'] for e in evs)
        mat_pay = 0.0
        if b == t['만기일']:
            mat_pay = state(b - timedelta(days=1))['미상환액면'] * t['만기상환율']
            end_carry = 0.0
        else:
            end_carry = comp(b)['장부금액']
        tags = []
        if b in {d for d, _ in flows}:
            tags.append('이자지급' if b != t['만기일'] else '만기')
        if b in (prior_end, fy_end):
            tags.append('결산')
        tags += [e['구분'] for e in evs]
        sched.append({'일자': b, '구분': '·'.join(tags), '기초장부금액': ca['장부금액'], '유효이자': intr,
                      '표면이자': paid, '상각액': intr - paid, '전환·상환 제거': off, '만기상환': mat_pay,
                      '기말장부금액': end_carry,
                      '미상환액면': 0.0 if b == t['만기일'] else comp(b)['미상환액면'],
                      '상환할증금': 0.0 if b == t['만기일'] else comp(b)['상환할증금'],
                      '전환권조정': 0.0 if b == t['만기일'] else comp(b)['전환권조정']})
    res['schedule'] = sched
    res['adj_schedule'] = adj_schedule(t, state, flows, eir, ev_res, comp, prior_end.month, fy_start, fy_end)

    # ③ 회계연도 집계
    if fy_end > t['발행일']:
        beg = comp(prior_end) if prior_end >= t['발행일'] else {kk: 0.0 for kk in comp(t['발행일'])}
        endc = comp(fy_end) if fy_end < t['만기일'] else {kk: 0.0 for kk in beg}
        fy_ev = [e for e in ev_res if fy_start <= e['일자'] <= fy_end]
        a0 = max(prior_end, t['발행일'])
        intr = E.interest_between(state, flows, eir, a0, min(fy_end, t['만기일']), ev_res)
        paid = E.coupons_paid_between(state, flows, a0, min(fy_end, t['만기일']), t['만기상환율'], t['만기일'])
        issued = split['최초장부금액'] if fy_start <= t['발행일'] <= fy_end else 0.0
        eq_issued = split['전환권대가'] if fy_start <= t['발행일'] <= fy_end else 0.0
        off = sum(e['부채제거'] for e in fy_ev)
        mat_pay = (state(t['만기일'] - timedelta(days=1))['미상환액면'] * t['만기상환율']
                   if fy_start <= t['만기일'] <= fy_end else 0.0)
        fy = {'기초': beg, '기말': endc, '발행': issued, '전환권대가_발행': eq_issued, '유효이자': intr,
              '표면이자': paid, '제거': off, '만기상환': mat_pay,
              '전환권대가_감소': sum(e['전환권대가감소'] for e in fy_ev),
              '전환주식수': sum(e.get('발행주식수', 0) for e in fy_ev),
              '상환손익': sum(e.get('상환손익', 0.0) for e in fy_ev), '이벤트': fy_ev}
        # 검증: 기초 + 발행 + 유효이자 − 표면이자 − 제거 − 만기상환 = 기말
        roll = beg['장부금액'] + issued + intr - paid - off - mat_pay
        fy['검증차이'] = roll - endc['장부금액']
        if abs(fy['검증차이']) > TOL_AMT:
            chk.append(('오류', f"장부금액 롤포워드 불일치 {fy['검증차이']:,.0f}원 — 계산 확인 필요"))
        res['fy'] = fy

        # 회사계상 대사
        rec = []
        for b in books_raw:
            bd = _date(b.get('기준일'))
            if not bd:
                continue
            ours = comp(bd)
            if bd == fy_end:
                ours = dict(ours, 당기이자비용=intr)
            for col, key in (('사채액면(원)', '미상환액면'), ('상환할증금(원)', '상환할증금'),
                             ('전환권조정(원)', '전환권조정'), ('장부금액(원)', '장부금액'),
                             ('당기이자비용(원)', '당기이자비용'), ('전환권대가(원)', '전환권대가')):
                cv = _num(b.get(col))
                if cv is None or key not in ours:
                    continue
                diff = cv - ours[key]
                flag = '' if abs(diff) <= max(TOL_AMT, abs(ours[key]) * TOL_RATE) else '확인'
                rec.append({'기준일': bd, '항목': col.replace('(원)', ''), '회사계상': cv,
                            '재계산': ours[key], '차이': diff, '판정': flag})
        res['book_rec'] = rec
        if any(r['판정'] for r in rec):
            chk.append(('확인', '회사계상 대비 재계산 차이 1% 초과 항목 있음 — 대사표 참고'))

        # 유동성: 당기말 이후 1년 이내 행사 가능한 Put 또는 만기
        horizon = date(fy_end.year + 1, fy_end.month, calendar.monthrange(fy_end.year + 1, fy_end.month)[1])
        nxt = [d for d, _ in t['puts'] if d > fy_end]
        past_put = any(d <= fy_end for d, _ in t['puts']) and nxt
        if endc['장부금액'] > 0:
            if t['만기일'] <= horizon:
                chk.append(('확인', f"만기 {t['만기일']} 가 결산일+1년 이내 → 유동부채 분류"))
            elif nxt and (nxt[0] <= horizon or past_put):
                chk.append(('확인', f"다음 조기상환청구일 {nxt[0]} (결산일+1년 이내) → 유동부채 분류 대상"
                                   " (K-IFRS 1001.69, 전환권 자본분류 시 전환은 결제로 보지 않음)"))
            else:
                chk.append(('정보', '결산일+1년 이내 조기상환청구일·만기 없음 → 비유동 분류 가능'))

    # 분류·조건 검토
    if t['리픽싱']:
        chk.append(('오류', '리픽싱 조항 있음 — 전환권이 fixed-for-fixed 요건 불충족 가능, 파생상품부채 분류 검토 '
                            '(본 앱 v1은 자본분류 회계처리만 지원)'))
    else:
        chk.append(('정보', '리픽싱 없음(입력 기준) — 전환권 자본분류(fixed-for-fixed) 전제'))
    for c in t['calls']:
        if '최대주주' in c['보유자']:
            chk.append(('확인', f"매도청구권 보유자 최대주주(행사가능 {c['행사가능비율']:.0%}) — 회사가 최대주주에게 "
                               '권리를 부여·양도했다면 별도 회계처리(자본거래/주식기준보상 등) 검토, 평가보고서상 '
                               '매도청구권 가치 확인'))
        elif '회사' in c['보유자']:
            chk.append(('확인', '매도청구권 보유자 회사 — 발행회사 콜옵션(내재파생상품) 주계약 밀접관련 여부 검토'))
    s = split
    # 검증: 발행가액 + 매도청구권 자산 = 부채 최초장부 + 전환권대가 + 거래비용
    gap = t['발행가액'] + s['매도청구권_자산'] - (s['최초장부금액'] + s['전환권대가'] + s['거래비용_부채'] + s['거래비용_자본'])
    if abs(gap) > 1:
        chk.append(('오류', f'발행가액 분리 합계 불일치 {gap:,.0f}'))
    return res


# ─── 출력 ────────────────────────────────────────────────────────────────────

THIN = Side(style='thin', color='BFBFBF')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD = PatternFill('solid', fgColor='D9E1F2')
TITLE = Font(bold=True, size=12)
BOLD = Font(bold=True)
FLAG_FILL = {'오류': PatternFill('solid', fgColor='FFC7CE'), '확인': PatternFill('solid', fgColor='FFEB9C')}


class Writer:
    def __init__(self, ws):
        self.ws, self.r = ws, 1

    def title(self, text):
        self.ws.cell(self.r, 1, text).font = TITLE
        self.r += 1

    def note(self, text):
        self.ws.cell(self.r, 1, text).font = Font(size=9, color='595959')
        self.r += 1

    def blank(self, n=1):
        self.r += n

    def table(self, headers, rows, fmts=None):
        fmts = fmts or {}
        for c, h in enumerate(headers, 1):
            cell = self.ws.cell(self.r, c, h)
            cell.font, cell.fill, cell.border = BOLD, HEAD, BORDER
            cell.alignment = Alignment(horizontal='center', wrap_text=True)
        self.r += 1
        for row in rows:
            for c, v in enumerate(row, 1):
                cell = self.ws.cell(self.r, c, v)
                cell.border = BORDER
                h = headers[c - 1]
                if h in fmts:
                    cell.number_format = fmts[h]
                elif isinstance(v, (date, datetime)):
                    cell.number_format = 'yyyy-mm-dd'
                elif isinstance(v, float) and abs(v) >= 1:
                    cell.number_format = '#,##0'
                if v in FLAG_FILL:
                    cell.fill = FLAG_FILL[v]
            self.r += 1

    def kv(self, pairs, fmts=None):
        fmts = fmts or {}
        for k, v in pairs:
            self.ws.cell(self.r, 1, k).font = BOLD
            cell = self.ws.cell(self.r, 2, v)
            if k in fmts:
                cell.number_format = fmts[k]
            elif isinstance(v, (date, datetime)):
                cell.number_format = 'yyyy-mm-dd'
            elif isinstance(v, float) and abs(v) >= 1:
                cell.number_format = '#,##0'
            self.r += 1


def widths(ws, w):
    for i, x in enumerate(w, 1):
        ws.column_dimensions[get_column_letter(i)].width = x


def save(results, out_path, company, fy_label, fy_start, fy_end, input_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = '요약'
    w = Writer(ws)
    w.title(f'전환사채 평가·회계처리 검증 — {company} ({fy_label})')
    w.note(f'검증기간: {fy_start} ~ {fy_end}  |  입력: {os.path.basename(input_path)}')
    if run_stamp:
        w.note(f'생성: {run_stamp.now_str()} / {run_stamp.pc_label()}')
    w.blank()
    rows = []
    for r in results:
        fy = r.get('fy')
        if not fy:
            continue
        rows.append([r['회차'], fy['기초']['장부금액'], fy['발행'], fy['유효이자'], fy['표면이자'],
                     fy['제거'], fy['만기상환'], fy['기말']['장부금액'], fy['기말']['미상환액면'],
                     fy['기말']['상환할증금'], fy['기말']['전환권조정'], fy['기말']['전환권대가'],
                     fy['전환주식수'], fy['상환손익'], fy['검증차이']])
    w.title('1. 회계연도 장부금액 변동 (재계산)')
    hdr = ['회차', '기초 장부금액', '발행', '이자비용(유효)', '표면이자 지급', '전환·상환 제거', '만기상환',
           '기말 장부금액', '기말 사채액면', '기말 상환할증금', '기말 전환권조정', '기말 전환권대가',
           '전환 발행주식수', '상환손익', '검증차이']
    if rows:
        tot = ['합계'] + [sum(x[i] for x in rows) for i in range(1, len(hdr))]
        rows.append(tot)
    w.table(hdr, rows, {'전환 발행주식수': '#,##0'})
    w.note('검증차이 = 기초 + 발행 + 이자비용 − 표면이자 − 제거 − 만기상환 − 기말 (0이어야 정상)')
    w.blank()

    w.title('2. 회사계상 대사')
    rec_rows = [[r['회차'], x['기준일'], x['항목'], x['회사계상'], x['재계산'], x['차이'], x['판정']]
                for r in results for x in r.get('book_rec', [])]
    if rec_rows:
        w.table(['회차', '기준일', '항목', '회사계상', '재계산', '차이', '판정'], rec_rows)
    else:
        w.note('회사계상 시트 입력 없음')
    w.blank()

    w.title('3. 공정가치 재계산 vs 외부평가')
    vrows = []
    for r in results:
        for v in r.get('valuation', []):
            for key in ('총가치', '일반사채', '조기상환청구권', '전환권', '매도청구권'):
                ext = v['외부'].get(key)
                diff = (v[key] - ext) if ext is not None else None
                rate = (diff / ext) if ext else None
                vrows.append([r['회차'], v['평가기준일'], key, v[key], ext, diff, rate])
    if vrows:
        w.table(['회차', '평가기준일', '구성요소', '재계산', '외부평가', '차이', '차이율'], vrows, {'차이율': '0.0%'})
    else:
        w.note('평가시점 입력 없음')
    w.blank()

    w.title('4. 검토사항')
    crow = [[r['회차'], lv, msg] for r in results for lv, msg in r['checks']]
    w.table(['회차', '구분', '내용'], crow)
    widths(ws, [12, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16, 12, 14, 12])
    ws.column_dimensions['C'].width = 22

    # 회차별 상세
    for r in results:
        t = r['terms']
        ws = wb.create_sheet(f"상각표_{r['회차']}"[:31])
        w = Writer(ws)
        w.title(f"{r['회차']} 발행조건·최초인식")
        w.kv([('발행일', t['발행일']), ('만기일', t['만기일']), ('액면총액', t['액면']), ('발행가액', t['발행가액']),
              ('거래비용', t['거래비용']), ('표면이자율', t['표면이자율']), ('이자지급주기(개월)', t['이자주기']),
              ('만기보장수익률', t['YTM']), ('만기상환율', t['만기상환율']), ('만기상환율 출처', t['만기상환율_출처']),
              ('전환가격', t['전환가격']), ('전환청구기간', f"{t['전환청구시작일']} ~ {t['전환청구종료일']}"),
              ('일반사채 시장이자율', t['시장이자율'])],
             {'표면이자율': '0.000%', '만기보장수익률': '0.000%', '만기상환율': '0.0000%', '일반사채 시장이자율': '0.000%'})
        if 'split' in r:
            s = r['split']
            w.blank()
            w.kv([('부채요소 측정', s['부채요소_출처']),
                  ('부채요소 공정가치', s['부채요소_공정가치']), ('매도청구권 자산', s['매도청구권_자산']),
                  ('자본요소(전환권) 공정가치', s['자본요소_공정가치']),
                  ('거래비용 — 부채', s['거래비용_부채']), ('거래비용 — 자본', s['거래비용_자본']),
                  ('최초 장부금액(부채)', s['최초장부금액']), ('전환권대가(자본)', s['전환권대가']),
                  ('유효이자율', s['유효이자율'])], {'유효이자율': '0.0000%'})
            w.note('잔여법(K-IFRS 1032.31): 자본요소 = 발행가액 − 부채요소(주계약+조기상환청구권) + 매도청구권 자산(자산인식 Y인 경우). '
                   '유효이자율은 부채 최초장부금액을 만기 계약현금흐름에 맞춰 산정.')
            sr = r.get('split_report')
            if sr:
                w.blank()
                w.title('최초인식 분리 — 재계산 vs 평가보고서')
                w.table(['구분', '재계산', '평가보고서', '차이', '차이율'],
                        [[nm, a, b, a - b, (a - b) / b if b else None] for nm, a, b in (
                            ('부채요소(주계약+조기상환권)', s['부채요소_공정가치'], sr['부채요소']),
                            ('매도청구권 자산', s['매도청구권_자산'], sr['매도청구권_자산']),
                            ('자본요소(전환권대가, 거래비용 전)', s['자본요소_공정가치'], sr['자본요소']))],
                        {'차이율': '0.0%'})
        if t['put_rows']:
            w.blank()
            w.title('조기상환청구권(Put)')
            w.table(['행사일', '조기상환수익률', '상환율'],
                    [[p['행사일'], p['조기상환수익률'], p['상환율']] for p in t['put_rows']],
                    {'조기상환수익률': '0.000%', '상환율': '0.0000%'})
        if t['call_rows']:
            w.blank()
            w.title('매도청구권(Call)')
            w.table(['행사시작일', '행사종료일', '지급기일', '행사가수익률', '행사가율', '적용행사가', '행사가능비율', '보유자',
                     '평가방법', '할인방식', '자산인식'],
                    [[c['행사시작일'], c['행사종료일'], c['지급기일'], c['행사가수익률'], c['행사가율'], c.get('_가격'),
                      c['행사가능비율'], c['보유자'], c['평가방법'], c['할인방식'], c['자산인식']] for c in t['call_rows']],
                    {'행사가수익률': '0.000%', '행사가율': '0.0000%', '적용행사가': '0.0000%', '행사가능비율': '0%'})
            if any(c.get('지급기일') for c in t['call_rows']):
                w.note('지급기일이 있는 행은 묶어서 버뮤다형(지급기일에만 행사)으로 평가')
        if r.get('schedule'):
            w.blank()
            w.title('유효이자율 상각표')
            hdr = ['일자', '구분', '기초장부금액', '유효이자', '표면이자', '상각액', '전환·상환 제거', '만기상환',
                   '기말장부금액', '미상환액면', '상환할증금', '전환권조정']
            w.table(hdr, [[x[h] for h in hdr] for x in r['schedule']])
        if r.get('events'):
            w.blank()
            w.title('전환·상환 처리')
            hdr = ['일자', '구분', '액면금액', '부채제거', '전환권대가감소', '발행주식수', '자본금', '주식발행초과금',
                   '지급액', '상환손익']
            w.table(hdr, [[e.get(h) for h in hdr] for e in r['events']], {'발행주식수': '#,##0'})
        widths(ws, [24, 18, 16, 16, 16, 16, 16, 16, 16, 16, 16, 16])

    # 전환권조정 분기별 상각표
    for r in results:
        rows = r.get('adj_schedule')
        if not rows:
            continue
        ws = wb.create_sheet(f"전환권조정상각_{r['회차']}"[:31])
        w = Writer(ws)
        w.title(f"{r['회차']} 전환권조정 상각표 (분기별, 유효이자율 {r['split']['유효이자율']:.4%})")
        w.note('전환권조정 = 사채액면 + 상환할증금 − 장부금액.  상각액 = 유효이자 − 표면이자 (이자비용 중 전환권조정 상각분).  '
               '상각기간: 만기 기준.  음영 = 당기 검증기간.')
        hdr = ['회계연도', '구분', '기간시작', '기간종료', '기초 전환권조정', '상각액', '전환·상환 감소', '기말 전환권조정',
               '유효이자', '표면이자', '사채액면', '상환할증금', '장부금액', '비고']
        num = ['기초 전환권조정', '상각액', '전환·상환 감소', '기말 전환권조정', '유효이자', '표면이자',
               '사채액면', '상환할증금', '장부금액']
        fmts = {h: '#,##0' for h in num}
        out, cur_rows = [], []
        for fy in sorted({x['회계연도'] for x in rows}):
            grp = [x for x in rows if x['회계연도'] == fy]
            for x in grp:
                out.append([x[h] for h in hdr])
                if x['당기']:
                    cur_rows.append(len(out))
            sub = [f'FY{fy} 소계', '', grp[0]['기간시작'], grp[-1]['기간종료'], grp[0]['기초 전환권조정'],
                   sum(x['상각액'] for x in grp), sum(x['전환·상환 감소'] for x in grp), grp[-1]['기말 전환권조정'],
                   sum(x['유효이자'] for x in grp), sum(x['표면이자'] for x in grp),
                   grp[-1]['사채액면'], grp[-1]['상환할증금'], grp[-1]['장부금액'], '']
            out.append(sub)
        tot_am = sum(x['상각액'] for x in rows)
        tot_dec = sum(x['전환·상환 감소'] for x in rows)
        out.append(['합계', '', rows[0]['기간시작'], rows[-1]['기간종료'], rows[0]['기초 전환권조정'], tot_am, tot_dec,
                    rows[-1]['기말 전환권조정'], sum(x['유효이자'] for x in rows), sum(x['표면이자'] for x in rows),
                    '', '', '', ''])
        top = w.r + 1
        w.table(hdr, out, fmts)
        cur_fill = PatternFill('solid', fgColor='FFF2CC')
        for i, row in enumerate(out):
            rr = top + i
            if str(row[0]).endswith('소계') or row[0] == '합계':
                for c in range(1, len(hdr) + 1):
                    ws.cell(rr, c).font = BOLD
            elif i + 1 in cur_rows:
                for c in range(1, len(hdr) + 1):
                    ws.cell(rr, c).fill = cur_fill
        w.blank()
        gap = rows[0]['기초 전환권조정'] - tot_am - tot_dec - rows[-1]['기말 전환권조정']
        w.kv([('최초 전환권조정', rows[0]['기초 전환권조정']), ('상각액 합계', tot_am),
              ('전환·상환 감소 합계', tot_dec), ('최종 잔액', rows[-1]['기말 전환권조정']),
              ('검증차이 (최초 − 상각 − 감소 − 최종)', gap)])
        widths(ws, [14, 8, 12, 12, 18, 16, 16, 18, 16, 14, 18, 16, 18, 24])

    # 공정가치·민감도
    ws = wb.create_sheet('공정가치')
    w = Writer(ws)
    w.title('평가시점별 공정가치 재계산 (TF 이항격자)')
    w.note('금리=기간구조이면 무위험·신용할인율 칸은 잔존만기 현물이자율(참고), 격자는 단계별 선도이자율로 할인.')
    w.note('총가치 = 전환사채(Put·전환 포함) − 매도청구권 가치.  전환권 = CB(콜 제외) − 일반사채(Put 포함).  '
           '조기상환청구권 = 일반사채(Put 포함) − 일반사채(Put 없음).  금액은 평가일 미상환액면 기준.')
    hdr = ['회차', '평가기준일', '미상환액면', '주가', '변동성', '무위험이자율', '신용할인율', '배당수익률', '금리',
           '총가치', '전환사채_콜제외', '일반사채_Put포함', '일반사채', '전환권', '조기상환청구권', '매도청구권']
    rows = []
    for r in results:
        for v in r.get('valuation', []):
            m = v['시장변수']
            T = E.yf(v['평가기준일'], r['terms']['만기일'])
            rr = [x.zero(T) if isinstance(x, E.Curve) else x for x in (m['무위험이자율'], m['신용위험조정할인율'])]
            rows.append([r['회차'], v['평가기준일'], v['미상환액면'], m['주가'], m['변동성'], rr[0], rr[1],
                         m['배당수익률'], v['금리']] + [v[h] for h in hdr[9:]])
    w.table(hdr, rows, {'변동성': '0.00%', '무위험이자율': '0.000%', '신용할인율': '0.000%', '배당수익률': '0.00%'})
    crow = []
    for r in results:
        for v in r.get('valuation', []):
            bd = v.get('매도청구권_할인방식별') or {}
            if bd:
                ext = v['외부'].get('매도청구권')
                crow.append([r['회차'], v['평가기준일']] + [bd.get(d) for d in E.CALL_DISCOUNTS] +
                            [ext, (bd.get('위험이자율', 0) - ext) / ext if ext else None])
    if crow:
        w.blank()
        w.title('매도청구권 — 할인방식별 재계산')
        w.table(['회차', '평가기준일'] + list(E.CALL_DISCOUNTS) + ['외부평가', '위험이자율 기준 차이율'], crow,
                {'위험이자율 기준 차이율': '0.0%'})
        w.note('위험이자율: 콜가치 전체를 신용위험 할인(평가기관 실무) / TF분할: 주식성분 무위험·채권성분 위험 할인 / '
               '무위험이자율: 전체 무위험 할인. 선택한 할인방식(매도청구Call 시트)이 총가치·회계처리에 반영됨')
    w.blank()
    w.title('민감도 (총가치)')
    srows = [[r['회차'], s['평가기준일'], s['시나리오'], s['총가치'], s['전환권'], s['총가치 변동'], s['변동률']]
             for r in results for s in r.get('sensitivity', [])]
    w.table(['회차', '평가기준일', '시나리오', '총가치', '전환권', '총가치 변동', '변동률'], srows, {'변동률': '0.0%'})
    widths(ws, [12, 12, 16, 12, 10, 12, 12, 10, 10, 16, 16, 16, 16, 16, 16, 16])

    if run_stamp:
        run_stamp.stamp_workbook(wb, f'전환사채 검증 {company} {fy_label}')
    os.makedirs(os.path.dirname(out_path), exist_ok=True)
    wb.save(out_path)


def main():
    ap = argparse.ArgumentParser(description='전환사채 전환권평가·회계처리 검증앱')
    ap.add_argument('company', nargs='?', default=None)
    ap.add_argument('--file', default=None, help='input_data/ 기준 입력 파일명')
    ap.add_argument('--fiscal-month', type=int, default=12)
    ap.add_argument('--fiscal-year', type=int, default=None, help='생략 시 파일명 fy 뒤 숫자')
    ap.add_argument('--interim-month', type=int, default=None, help='중간결산월(예: 9 → 9월말 기준)')
    ap.add_argument('--steps-per-year', type=int, default=252, help='격자 연간 단계 수(기본 252)')
    a = ap.parse_args()

    path = find_input(a.company, a.file)
    m = re.search(r'convertible_(.+?)_information_fy(\d{4})', os.path.basename(path))
    company = a.company or (m.group(1) if m else 'company')
    fy = a.fiscal_year or (int(m.group(2)) if m else date.today().year)
    fy_start, fy_end, prior_end = fy_bounds(fy, a.fiscal_month, a.interim_month)
    print(f'[{company}] 전환사채 검증 — FY{fy} ({fy_start} ~ {fy_end})')
    print(f'  입력: {os.path.basename(path)}')

    terms = build_terms(path)
    if not terms:
        sys.exit('[오류] 발행조건 시트에 유효한 회차가 없습니다(예시 행은 무시됨).')
    ev = group_by(_read_sheet(path, '변동내역'))
    vals = group_by(_read_sheet(path, '평가시점'))
    books = group_by(_read_sheet(path, '회사계상'))
    curves = load_curves(path)
    if curves:
        print(f'  금리기간구조: {len(curves)}개 평가시점')

    results = []
    for k, t in terms.items():
        print(f'  - {k}: 액면 {t["액면"]:,.0f}, 만기상환율 {t["만기상환율"]:.4%}, '
              f'Put {len(t["puts"])}건, Call {len(t["calls"])}건')
        results.append(process(t, ev.get(k, []), vals.get(k, []), books.get(k, []),
                               fy_start, fy_end, prior_end, a.steps_per_year, curves))
        for lv, msg in results[-1]['checks']:
            if lv != '정보':
                print(f'      [{lv}] {msg}')

    suffix = f'_{a.interim_month:02d}월' if a.interim_month else ''
    out = os.path.join(OUTPUT_DIR, f'convertible_schedule_{company}_{fy}{suffix}.xlsx')
    save(results, out, company, f'FY{fy}{suffix}', fy_start, fy_end, path)
    print(f'[완료] {out}')


if __name__ == '__main__':
    main()
