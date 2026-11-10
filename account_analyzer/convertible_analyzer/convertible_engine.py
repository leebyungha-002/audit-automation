# -*- coding: utf-8 -*-
"""전환사채(CB) 평가·회계처리 계산 엔진 (입출력 없음, 순수 계산).

1) 공정가치: Tsiveriotis-Fernandes(TF) 이항격자
   - 주가 CRR 격자, 현금부분(B)은 신용위험조정할인율, 주식부분(V-B)은 무위험이자율로 할인
   - 조기상환청구권(Put, 투자자): 행사일에 V = max(V, 상환가액) → 현금
   - 매도청구권(Call, 최대주주/회사): 행사기간 중 V > max(행사가, 전환가치)이면 콜 행사
       → 투자자는 전환가치가 더 크면 전환, 아니면 행사가 수령
     행사가능비율 f: 전체가치 = (1-f)·V(콜없음) + f·V(콜100%)  (콜은 일부 물량에만 적용)
   - 전환: 전환청구기간 중 V = max(V, 전환비율×주가) → 주식
2) 회계처리(K-IFRS 1032 복합금융상품, 전환권 자본분류)
   - 부채요소 = 계약상 현금흐름(표면이자+만기상환액)을 일반사채 시장이자율로 할인
   - 자본요소(전환권대가) = 발행가액 − 부채요소, 거래비용은 공정가치 비율로 안분
   - 유효이자율: 최초 장부금액 = 현금흐름 현재가치가 되는 연복리 이자율
   - 임의 시점 장부금액 = 잔여 현금흐름을 유효이자율로 할인 × 미상환 비율

금리·수익률 입력은 모두 연복리(%)이며 격자에서는 연속복리 ln(1+r)로 변환한다.
기간(년) = 일수 / 365.
"""
import math
from datetime import date, timedelta

import numpy as np


# ─── 날짜·현금흐름 ───────────────────────────────────────────────────────────

def yf(d0: date, d1: date) -> float:
    return (d1 - d0).days / 365.0


def add_months(d: date, months: int) -> date:
    y, m = divmod(d.month - 1 + months, 12)
    y += d.year
    m += 1
    import calendar
    last = calendar.monthrange(y, m)[1]
    return date(y, m, min(d.day, last))


def coupon_dates(issue: date, maturity: date, freq_months: int) -> list:
    """발행일부터 freq_months 간격 이자지급일(만기일 포함)."""
    if not freq_months or freq_months <= 0:
        return [maturity]
    out, k = [], 1
    while True:
        d = add_months(issue, freq_months * k)
        if d >= maturity:
            break
        out.append(d)
        k += 1
    out.append(maturity)
    return out


def contract_cashflows(issue, maturity, coupon_rate, freq_months, redemption_ratio):
    """액면 1 기준 계약상 현금흐름 [(일자, 금액)] — 표면이자 + 만기상환액."""
    cds = coupon_dates(issue, maturity, freq_months)
    cpn = coupon_rate * (freq_months / 12.0) if coupon_rate else 0.0
    flows = [(d, cpn) for d in cds]
    flows[-1] = (maturity, flows[-1][1] + redemption_ratio)
    return [(d, a) for d, a in flows if a]


def redemption_ratio_from_yield(issue, at_date, coupon_rate, freq_months, y):
    """보장수익률 y(연복리)를 맞추는 상환율(액면 1 기준, 해당일 표면이자 제외).

    상환율 = (1+y)^T − Σ 표면이자_i × (1+y)^(T − t_i)   (t_i ≤ T 인 기지급 이자)
    """
    T = yf(issue, at_date)
    val = (1 + y) ** T
    if coupon_rate:
        cpn = coupon_rate * freq_months / 12.0
        for d in coupon_dates(issue, at_date, freq_months):
            if d <= at_date:
                val -= cpn * (1 + y) ** yf(d, at_date)
    # 해당일에 지급되는 표면이자는 별도로 지급되므로 상환율에서 제외된 상태(위 루프에 포함돼 차감됨)
    return val


def pv(flows, at: date, y: float) -> float:
    return sum(a / (1 + y) ** yf(at, d) for d, a in flows if d > at)


def solve_rate(flows, at: date, target: float, lo=-0.5, hi=1.0) -> float:
    """pv(flows, at, r) = target 인 r (이분법)."""
    f_lo, f_hi = pv(flows, at, lo) - target, pv(flows, at, hi) - target
    if f_lo * f_hi > 0:
        raise ValueError('유효이자율 계산 실패: 현금흐름과 장부금액 확인')
    for _ in range(200):
        mid = (lo + hi) / 2
        f_mid = pv(flows, at, mid) - target
        if abs(f_mid) < 1e-12:
            break
        if f_lo * f_mid <= 0:
            hi = mid
        else:
            lo, f_lo = mid, f_mid
    return (lo + hi) / 2


# ─── TF 이항격자 ─────────────────────────────────────────────────────────────

def tf_lattice(S0, sigma, r_f, r_c, q, val_date, maturity, conv_ratio, redemption,
               coupons, conv_window, puts=(), call_window=None, call_price_fn=None,
               steps_per_year=252, min_steps=200, allow_conversion=True):
    """액면 1 기준 CB 가치.

    coupons     : [(일자, 금액)] (val_date 이후분만 사용)
    conv_window : (시작일, 종료일) 전환청구기간
    puts        : [(일자, 상환율)]
    call_window : (시작일, 종료일) or None,  call_price_fn(일자) → 행사가(액면 1 기준)
    r_f, r_c, q : 연복리 → 내부에서 연속복리 변환
    반환: (총가치 V, 현금부분 B)
    """
    T = yf(val_date, maturity)
    if T <= 0:
        return redemption, redemption
    N = max(min_steps, int(math.ceil(T * steps_per_year)))
    dt = T / N
    rf, rc, qq = math.log(1 + r_f), math.log(1 + r_c), math.log(1 + q)
    u = math.exp(sigma * math.sqrt(dt))
    d = 1 / u
    p = (math.exp((rf - qq) * dt) - d) / (u - d)
    if not 0 < p < 1:
        raise ValueError(f'격자 확률 범위 오류(p={p:.4f}) — 단계 수·변동성 확인')
    disc_f, disc_c = math.exp(-rf * dt), math.exp(-rc * dt)

    def step_of(dd):
        return int(round(yf(val_date, dd) / dt))

    cpn = np.zeros(N + 1)
    for dd, a in coupons:
        if val_date < dd <= maturity:
            cpn[min(N, step_of(dd))] += a
    put_at = {}
    for dd, pr in puts:
        if val_date < dd < maturity:
            put_at[step_of(dd)] = pr
    c0 = step_of(conv_window[0]) if conv_window and conv_window[0] else 0
    c1 = step_of(conv_window[1]) if conv_window and conv_window[1] else N
    if call_window and call_price_fn:
        k0 = max(0, step_of(call_window[0]))
        k1 = min(N, step_of(call_window[1]))
    else:
        k0, k1 = 1, 0

    def S_at(i):
        j = np.arange(i + 1)
        return S0 * u ** j * d ** (i - j)

    # 만기
    conv = conv_ratio * S_at(N)
    V = np.full(N + 1, redemption + cpn[N])
    B = V.copy()
    if allow_conversion and c0 <= N <= c1 + 1:
        m = conv > V
        V = np.where(m, conv, V)
        B = np.where(m, 0.0, B)

    for i in range(N - 1, -1, -1):
        Bh = disc_c * (p * B[1:] + (1 - p) * B[:-1])
        Eh = disc_f * (p * (V[1:] - B[1:]) + (1 - p) * (V[:-1] - B[:-1]))
        Vh = Bh + Eh
        if cpn[i]:
            Vh = Vh + cpn[i]
            Bh = Bh + cpn[i]
        V, B = Vh, Bh
        conv = conv_ratio * S_at(i)
        can_conv = allow_conversion and c0 <= i <= c1
        if k0 <= i <= k1 and i > 0:
            cp = call_price_fn(val_date + timedelta(days=i * dt * 365))
            cap = np.maximum(cp, conv) if can_conv else np.full_like(V, cp)
            called = V > cap
            if can_conv:
                to_conv = called & (conv >= cp)
                to_cash = called & ~(conv >= cp)
                V = np.where(to_conv, conv, np.where(to_cash, cp, V))
                B = np.where(to_conv, 0.0, np.where(to_cash, cp, B))
            else:
                V = np.where(called, cp, V)
                B = np.where(called, cp, B)
        if i in put_at:
            pr = put_at[i]
            m = pr > V
            V = np.where(m, pr, V)
            B = np.where(m, pr, B)
        if can_conv:
            m = conv > V
            V = np.where(m, conv, V)
            B = np.where(m, 0.0, B)
    return float(V[0]), float(B[0])


def value_components(terms: dict, mkt: dict, val_date: date, steps_per_year=252):
    """평가시점 1건의 구성요소별 공정가치(액면 1 기준).

    반환 dict:
      총가치        : CB 전체(투자자 관점, 콜 행사가능비율 반영)
      일반사채_Put포함 : 전환권 없음, Put 있음
      일반사채       : 전환권·Put 없음
      전환권        : (Put 있는 CB, 콜 없음) − 일반사채_Put포함
      조기상환청구권  : 일반사채_Put포함 − 일반사채
      매도청구권     : 콜 보유자 가치 = f × (V콜없음 − V콜100%)
    """
    common = dict(S0=mkt['주가'], sigma=mkt['변동성'], r_f=mkt['무위험이자율'],
                  r_c=mkt['신용위험조정할인율'], q=mkt.get('배당수익률', 0.0),
                  val_date=val_date, maturity=terms['만기일'], conv_ratio=terms['전환비율'],
                  redemption=terms['만기상환율'], coupons=terms['현금흐름_이자'],
                  conv_window=(terms['전환청구시작일'], terms['전환청구종료일']),
                  steps_per_year=steps_per_year)
    puts = terms.get('puts', [])
    calls = terms.get('calls', [])

    v_nocall, _ = tf_lattice(puts=puts, **common)
    straight_put, _ = tf_lattice(puts=puts, allow_conversion=False, **common)
    straight, _ = tf_lattice(puts=(), allow_conversion=False, **common)

    call_val = 0.0
    total = v_nocall
    for c in calls:  # 콜이 여러 행이면 각 행사기간·비율별로 순차 반영(통상 1건)
        f = c['행사가능비율']
        if f <= 0:
            continue
        v_call, _ = tf_lattice(puts=puts, call_window=(c['시작일'], c['종료일']),
                               call_price_fn=c['가격함수'], **common)
        call_val += f * (v_nocall - v_call)
        total -= f * (v_nocall - v_call)
    return {
        '총가치': total,
        '전환사채_콜제외': v_nocall,
        '일반사채_Put포함': straight_put,
        '일반사채': straight,
        '전환권': v_nocall - straight_put,
        '조기상환청구권': straight_put - straight,
        '매도청구권': call_val,
    }


# ─── 회계처리 ────────────────────────────────────────────────────────────────

def initial_split(issue_price, face, flows_unit, issue_date, market_rate, txn_cost):
    """발행 시 부채·자본 분리(총액 기준)."""
    liab_fv = face * pv(flows_unit, issue_date, market_rate)
    equity_fv = issue_price - liab_fv
    cost_l = txn_cost * liab_fv / issue_price if issue_price else 0.0
    cost_e = txn_cost - cost_l
    carrying0 = liab_fv - cost_l
    eir = solve_rate(flows_unit, issue_date, carrying0 / face)
    return {'부채요소_공정가치': liab_fv, '자본요소_공정가치': equity_fv,
            '거래비용_부채': cost_l, '거래비용_자본': cost_e,
            '최초장부금액': carrying0, '전환권대가': equity_fv - cost_e, '유효이자율': eir}


def carrying_unit(flows_unit, at: date, eir: float) -> float:
    """액면 1 기준 장부금액(at 시점, at 당일 지급분 제외)."""
    return pv(flows_unit, at, eir)


def run_ledger(face, flows_unit, issue_date, eir, equity0, events, conv_price, par_per_share,
               until: date):
    """발행일~until 까지 이벤트(전환·상환) 반영 원장.

    events: [{'일자', '구분'(전환/조기상환/매입소각), '액면금액', '지급액'}] 일자순
    반환: (잔액함수 state_at(d), 이벤트 결과 리스트)
    state_at(d) → {'미상환액면', '장부금액', '전환권대가'}
    """
    out = face
    eq = equity0
    history = [(issue_date, out, eq)]   # (일자, 이후 미상환액면, 이후 전환권대가)
    results = []
    for ev in sorted(events, key=lambda e: e['일자']):
        if ev['일자'] > until:
            break
        x = min(ev['액면금액'], out)
        if x <= 0:
            continue
        ratio = x / out
        carry_before = out * carrying_unit(flows_unit, ev['일자'], eir)
        liab_off = carry_before * ratio
        eq_off = eq * ratio
        r = {'일자': ev['일자'], '구분': ev['구분'], '액면금액': x, '부채제거': liab_off,
             '전환권대가감소': eq_off}
        if ev['구분'] == '전환':
            shares = math.floor(x / conv_price) if conv_price else 0
            capital = shares * par_per_share
            r.update({'발행주식수': shares, '자본금': capital,
                      '주식발행초과금': liab_off + eq_off - capital})
        else:
            paid = ev.get('지급액') or 0.0
            r.update({'지급액': paid, '상환손익': liab_off - paid})
        out -= x
        eq -= eq_off
        history.append((ev['일자'], out, eq))
        results.append(r)

    def state_at(dd: date):
        o, e = face, equity0
        for hd, ho, he in history:
            if hd <= dd:
                o, e = ho, he
        if dd < issue_date:
            return {'미상환액면': 0.0, '장부금액': 0.0, '전환권대가': 0.0}
        return {'미상환액면': o, '장부금액': o * carrying_unit(flows_unit, dd, eir), '전환권대가': e}

    return state_at, results


def interest_between(face_state, flows_unit, eir, d0: date, d1: date, events_done) -> float:
    """(d0, d1] 유효이자비용 = Σ 구간별 [장부금액 증가 + 지급 표면이자], 이벤트 시점에 끊어 계산."""
    cuts = sorted({d0, d1} | {e['일자'] for e in events_done if d0 < e['일자'] < d1})
    total = 0.0
    for a, b in zip(cuts[:-1], cuts[1:]):
        st = face_state(a)
        o = st['미상환액면']
        if o <= 0:
            continue
        ca = carrying_unit(flows_unit, a, eir)
        cb = carrying_unit(flows_unit, b, eir)
        paid = sum(amt for dd, amt in flows_unit if a < dd <= b)
        total += o * (cb - ca + paid)
    return total


def coupons_paid_between(face_state, flows_unit, d0, d1, redemption_ratio, maturity):
    """(d0, d1] 지급 표면이자(만기상환원금·할증금 제외)."""
    tot = 0.0
    for dd, amt in flows_unit:
        if d0 < dd <= d1:
            a = amt - (redemption_ratio if dd == maturity else 0.0)
            tot += face_state(dd - timedelta(days=1))['미상환액면'] * a
    return tot
