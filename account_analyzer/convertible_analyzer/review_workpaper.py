"""전환사채 평가인 리뷰 코멘트 검토 조서 생성 (graphy 6회차, Review Comment – 평가인 #2 반영).

실행: python review_workpaper.py
출력: output/조서_CB평가_리뷰코멘트_검토_graphy_6회차_v2.xlsx

재계산 조건(평가보고서와 같은 조건): 주간 격자(연 52단계), 보고서 p.25 KIS 사모무보증 B-·국고채 매트릭스
정상 보간(분기 부트스트래핑), 변동성 106.15%(평가인 회신: 연환산 250일).
매도청구권은 리뷰어 수정 요청 방식 — T&F W/WO, MIN[Call, MAX(만기상환, Put, Con, HV)],
콜 행사기간 중 전환 불가, Call 노드는 무위험이자율 할인.
"""
import os
import sys
from datetime import date

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
import convertible_engine as E  # noqa: E402
import convertible_schedule as S  # noqa: E402

INPUT = os.path.join(BASE, 'input_data', 'convertible_graphy_information_fy2026.xlsx')
OUT = os.path.join(BASE, 'output', '조서_CB평가_리뷰코멘트_검토_graphy_6회차_v2.xlsx')

VD = date(2026, 7, 21)
FACE = 48_300_000_000
ISSUE_PRICE = 48_300_000_000
STEPS = 52
SIGMA = 1.0615
SIGMA_252 = 1.0615 * (250 / 252) ** 0.5
BLOCK = (date(2027, 7, 21), date(2029, 1, 21))   # 콜 행사기간 — 콜 대상 물량 전환 불가
TEN = [x[1] for x in S.TENOR_COLS]
RF = [2.774, 3.005, 3.279, 3.36, 3.537, 3.69, 3.805, 3.865, 4.01, 4.116, 4.233, 4.33]
RD = [13.242, 14.56, 15.729, 16.502, 17.937, 19.342, 20.245, 21.076, 21.988, 22.663, 23.617, 24.711]
REP = {'일반사채': 14_978_056_628, '조기상환청구권': 13_979_156_617, '전환권': 33_113_825_847,
       '매도청구권': 8_447_328_620, '총가치': 53_623_710_472}


def curve(ys, shift=False):
    y = [v / 100 for v in ys]
    if shift:   # 리뷰어 지적 오류 재현: 만기 구간이 한 칸씩 밀려 각 만기에 다음 만기 YTM 적용
        y = y[1:] + [y[-1]]
    return E.Curve(TEN, y)


def components(t, sigma=SIGMA, shift=False):
    mkt = {'주가': 17650, '변동성': sigma, '무위험이자율': curve(RF, shift),
           '신용위험조정할인율': curve(RD, shift), '배당수익률': 0.0}
    r = E.value_components(t, mkt, VD, STEPS)
    out = {k: r[k] * FACE for k in ('일반사채', '조기상환청구권', '전환권', '전환사채_콜제외', '매도청구권')}
    out['콜_Compound'] = {k: v * FACE for k, v in r['매도청구권_할인방식별'].items()}
    common = dict(S0=17650, sigma=sigma, r_f=mkt['무위험이자율'], r_c=mkt['신용위험조정할인율'], q=0.0,
                  val_date=VD, maturity=t['만기일'], conv_ratio=t['전환비율'], redemption=t['만기상환율'],
                  coupons=t['현금흐름_이자'], conv_window=(t['전환청구시작일'], t['전환청구종료일']),
                  steps_per_year=STEPS, puts=t['puts'])
    call = t['calls'][0]
    sched, f = dict(call['행사일정']), call['행사가능비율']
    v_wo, _ = E.tf_lattice(**common)
    v_wo_blk, _ = E.tf_lattice(no_conv=BLOCK, **common)
    ww = {}
    for cash in ('무위험이자율', '위험이자율'):
        for blk in (True, False):
            v_w, _ = E.tf_lattice(call_steps=sched, call_cash=cash, no_conv=BLOCK if blk else None, **common)
            ww[(cash, blk)] = f * (v_wo - v_w) * FACE
    out['콜_WW'] = ww
    out['V_wo'] = v_wo * FACE
    out['V_wo_전환제한'] = v_wo_blk * FACE
    out['콜'] = ww[('무위험이자율', True)]
    out['총가치'] = out['전환사채_콜제외'] - out['콜']
    return out


# ─── 서식 ────────────────────────────────────────────────────────────────────
THIN = Side(style='thin', color='999999')
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HEAD = PatternFill('solid', fgColor='DDEBF7')
TITLE = Font(bold=True, size=13)
BOLD = Font(bold=True)
WRAP = Alignment(wrap_text=True, vertical='top')
WON, PCT, PCT3 = '#,##0;[Red]-#,##0', '0.00%;[Red]-0.00%', '0.000%'


class Sheet:
    def __init__(self, ws, widths):
        self.ws, self.r = ws, 1
        for i, w in enumerate(widths):
            ws.column_dimensions[chr(65 + i)].width = w

    def title(self, text):
        self.ws.cell(self.r, 1, text).font = TITLE
        self.r += 2

    def sub(self, text):
        self.ws.cell(self.r, 1, text).font = BOLD
        self.r += 1

    def note(self, text):
        c = self.ws.cell(self.r, 1, text)
        c.font = Font(italic=True, color='555555')
        self.r += 1

    def blank(self, n=1):
        self.r += n

    def table(self, headers, rows, fmts=None):
        for j, h in enumerate(headers, 1):
            c = self.ws.cell(self.r, j, h)
            c.font, c.fill, c.border, c.alignment = BOLD, HEAD, BOX, Alignment(wrap_text=True, vertical='center')
        self.r += 1
        start = self.r
        for row in rows:
            for j, v in enumerate(row, 1):
                c = self.ws.cell(self.r, j, v)
                c.border, c.alignment = BOX, WRAP
                if fmts and j - 1 < len(fmts) and fmts[j - 1]:
                    c.number_format = fmts[j - 1]
            self.r += 1
        return start


def m(v):
    return f'{v / 1e8:,.1f}억원'


def build():
    t = S.build_terms(INPUT)['6회차']
    base = components(t)
    shifted = components(t, shift=True)
    v252 = components(t, sigma=SIGMA_252)
    call_rf, call_rd = base['콜_WW'][('무위험이자율', True)], base['콜_Compound']['위험이자율']
    d_call = call_rf - REP['매도청구권']
    liab_rep = REP['일반사채'] + REP['조기상환청구권']
    liab_new = base['일반사채'] + base['조기상환청구권']
    eq_rep = ISSUE_PRICE - liab_rep + REP['매도청구권']
    eq_new = ISSUE_PRICE - liab_new + call_rf

    wb = Workbook()

    # ── 1. 검토개요 ──
    s = Sheet(wb.active, [16, 120])
    wb.active.title = '검토개요'
    s.title('graphy(T사) 제6회 전환사채 공정가치 평가 — 평가인 리뷰 코멘트 검토 조서 (#2 회신 반영)')
    s.table(['항목', '내용'], [
        ['대상', '제6회 무보증 사모 전환사채 (발행일 2026-07-21, 만기 2031-07-21, 액면·발행가액 48,300,000,000원)'],
        ['검토 자료 ①', 'Review Comment – 평가인 #1 (Reviewed By Samduk) — 주가변동성 [1], Bootstrapping [2], 콜옵션 [3] 질의'],
        ['검토 자료 ②', 'Review Comment – 평가인 #2 — 평가인 회신(평가상세파일 1~5번 근거)과 리뷰어 결론(OK), '
                      '추가 지적 2건: [2] YTM 보간 오류, [3] 콜옵션 평가방법 W/WO 수정 요청'],
        ['검토 자료 ③', 'CB6 공정가치 평가보고서(A회계법인, 평가기준일 2026-07-21 최초발행)'],
        ['검토 자료 ④', 'convertible_analyzer 독립 재계산 — TF 이항격자 주간 52단계(보고서 262노드와 같은 조건), '
                      '보고서 p.25 KIS 사모무보증 B-·국고채 매트릭스 정상 보간 부트스트래핑, 변동성 106.15%. '
                      '콜옵션은 리뷰어 요청 방식(W/WO) 구현(review_workpaper.py)'],
        ['검토 목적', '평가인 회신에 대한 리뷰어 결론을 반영해 종결 항목과 미결 항목(평가인 수정 요청)을 나누고, '
                    '리뷰어 수정 요청대로 재계산했을 때 평가금액과 최초인식 회계처리에 미치는 영향을 정리'],
        ['평가인 회신', 'Review Comment – 평가인 #2 수령(2026-11-27)'],
        ['작성', '(작성자 / 일자)'],
        ['검토', '(검토자 / 일자)'],
    ])
    s.blank()
    s.sub('검토 결과 요약')
    s.table(['구분', '내용'], [
        ['종결', '[1] 주가변동성 — 평가인 회신: 연환산 250일 적용(평가상세파일 4번 4번 sheet). 리뷰어 OK. '
               '106.15%는 250일 기준 연변동성으로 확인됨. 참고로 감사팀 격자에서 105.73%를 쓰면 보고서 전환권이 '
               '−0.001%로 맞지만 격자 단계(dt) 설정 차이 수준이고 106.15% 적용 시 전환권 차이도 +0.68%라 수용'],
        ['종결', '[2]-1 신용등급 B- — NICE평가정보 신용평가보고서(2024년말 기준). 유효기간은 지났으나 평가기준일 현재 '
               '등급이 CCC+ 이하로 판단돼 B- 적용 수용(리뷰어 OK)'],
        ['종결', '[2]-2 YTM Curve — KIS-NET 공시 YTM 기준, 일별 금리 산정 후 Bootstrapping(평가상세파일 2·3번). '
               '공모 배율 적용 확인(리뷰어 OK)'],
        ['미결(수정요청)', f'[2]-3 YTM 보간 오류(리뷰어 추가 지적) — "01_YTM보간법"·"01_YTM보간법(Rf)" sheet 구간1 끝(2026-10-21)에 '
                     f'6개월 YTM(13.914%) 대신 3개월 YTM(12.438%)을 써야 함 → 주계약·풋옵션 과대할인. '
                     f'정상 보간 재계산 주계약 {base["일반사채"]:,.0f}원(보고서 대비 {base["일반사채"] - REP["일반사채"]:+,.0f}원). '
                     f'같은 유형 오류를 p.25 매트릭스에 재현하면 주계약이 {shifted["일반사채"] - base["일반사채"]:+,.0f}원 바뀌어 영향이 클 수 있음'],
        ['미결(수정요청)', f'[3] 콜옵션 — 평가인 방식은 차액정산형 접근(MAX(V−EXE, 0) vs 시간가치를 rd로 할인)으로 판단됨. '
                     f'본 콜옵션은 총액교환형이므로 W/WO, MIN[Call, MAX(만기상환, Put, Con, HV)], 콜 행사기간 중 전환권 0, '
                     f'Call 노드 rf 할인으로 수정 요청. 감사팀 재계산 {call_rf:,.0f}원 — 보고서 {REP["매도청구권"]:,.0f}원 대비 '
                     f'{d_call:+,.0f}원({d_call / REP["매도청구권"]:+.1%})'],
        ['회계 영향', f'리뷰 반영 시 최초인식 자본요소(전환권대가) {eq_rep:,.0f}원 → {eq_new:,.0f}원({eq_new - eq_rep:+,.0f}원), '
                    f'부채요소(주계약+조기상환권) {liab_new - liab_rep:+,.0f}원, 매도청구권 자산 {d_call:+,.0f}원. '
                    '평가인 수정 보고서 수령 후 확정(재계산결과 시트)'],
        ['기타', '보고서 기재 오류(p.26 조기상환권 행사가능기간 12개월, p.27 "유사기업의 주가 변동성" 표제)는 '
               '수정 보고서에서 함께 정정 요청'],
    ])

    # ── 2. 리뷰코멘트대응 ──
    s = Sheet(wb.create_sheet('리뷰코멘트대응'), [8, 34, 40, 34, 52, 34, 14])
    s.title('리뷰 코멘트별 검토 — 평가인 회신(#2)·리뷰어 결론·감사팀 재계산')
    s.table(['No', '리뷰 코멘트(#1)', '평가인 회신(#2)', '리뷰어 결론(#2)', '감사팀 재계산·분석', '후속 조치', '상태'], [
        ['[1]', '주가변동성: 180영업일 일별 변동률 → 연간 변동성 환산에 적용한 영업일 수는?',
         '연환산 일수 250일 적용. 산정내역은 평가상세파일 4번파일 4번 sheet',
         'OK',
         f'106.15%를 250일 기준 연변동성으로 보고 재계산: 전환권 {base["전환권"]:,.0f}원(보고서 대비 '
         f'{base["전환권"] / REP["전환권"] - 1:+.2%}). 105.73%(=106.15%×√(250/252)) 적용 시 '
         f'{v252["전환권"]:,.0f}원({v252["전환권"] / REP["전환권"] - 1:+.3%}) — 격자 dt 설정 차이 수준으로 판단',
         '없음', '종결'],
        ['[2]-1', '회사 신용등급 B-의 출처? 출처와 Screen Copy 요청',
         'NICE평가정보 신용평가보고서(2024년말 기준) B-. 화면캡쳐는 평가상세파일 4번파일 1번 sheet',
         'OK — 유효기간이 지났지만 평가기준일 현재 등급이 CCC+ 이하라 B- 적용 수용',
         '재계산은 B- 매트릭스 사용. 리뷰어 결론에 동의',
         '평가상세파일 4번 1번 sheet 화면캡쳐 조서 첨부', '종결'],
        ['[2]-2', 'YTM Curve 출처는 KIS-NET — B- YTM Curve를 어떻게 산정했는지?',
         'KIS-NET 공시 YTM 기준(평가상세파일 4번 2번 sheet). 일별 금리 산정 후 bootstrapping으로 현행·선도이자율 '
         '산정(2번파일 T-F, 3번파일 B-D-T)',
         'OK — 공모의 배율 적용',
         '보고서 p.7·23 대표 위험이자율 24.66%가 p.25 KIS 5Y 22.663%와 2.000%p 차이 났던 이슈는 '
         '공모 배율 적용 및 [2]-3 보간 오류와 관련된 것으로 보임 → [2]-3 수정 보고서에서 함께 확인',
         '[2]-3과 함께 수정 보고서로 확인', '종결(리뷰어 OK)'],
        ['[2]-3\n(추가)', '리뷰어 추가 지적: "01_YTM보간법" sheet 구간1 끝 2026-10-21 YTM은 13.914%(6개월)가 아니라 '
                        '12.438%(3개월)가 돼야 함. "01_YTM보간법(Rf)" sheet도 같은 오류 → 주계약·풋옵션 과대할인',
         '(회신 대기)', '수정 요청',
         f'① 정상 보간 재계산(p.25 매트릭스): 주계약 {base["일반사채"]:,.0f}원(보고서 대비 '
         f'{base["일반사채"] - REP["일반사채"]:+,.0f}원), 조기상환권 {base["조기상환청구권"]:,.0f}원'
         f'({base["조기상환청구권"] - REP["조기상환청구권"]:+,.0f}원)\n'
         f'② 오류 유형 영향: p.25 매트릭스에서 만기 구간을 한 칸씩 밀면 주계약 {shifted["일반사채"]:,.0f}원'
         f'(정상 대비 {shifted["일반사채"] / base["일반사채"] - 1:+.1%}) — 오류가 전 구간에 있으면 금액 영향이 큼\n'
         '③ 리뷰어가 인용한 평가인 파일 값(3M 12.438%, 6M 13.914%)은 보고서 p.25 매트릭스(3M 13.242%, 6M 14.560%)와 '
         '다름 → 공모 배율 적용 후 값으로 보이나, 보고서 기재 매트릭스와 실제 적용 매트릭스가 다르므로 평가인 확인 필요',
         '보간 수정 후 재평가 보고서·평가상세파일(2·3번) 수령 → 재계산과 대사. 보고서 p.25에 실제 적용 매트릭스(배율 반영) 기재 요청',
         '미결'],
        ['[3]-1~7', '콜옵션 평가방식(W/WO·Compound), 수식, Call 노드 할인율(RF/RD), Compound 과정·할인율, '
                    '전환불가·콜 우선 조항 반영, 평가 엑셀 파일 요청',
         '제3자 지정가능 콜옵션으로 별도 T-F 모형 평가(평가상세파일 5번). W/WO 미적용. 할인은 위험이자율. '
         '전환불가·콜 우선 조항은 T-F 모형 노드별로 자동 검토된다고 보고 별도 고려하지 않음',
         'OK — [3]-8로 통합 질의',
         f'감사팀 Compound(별도격자)+콜가치 전체 RD 할인이 보고서를 재현: {call_rd:,.0f}원'
         f'({call_rd / REP["매도청구권"] - 1:+.2%}) → 평가인 방식에 대한 리뷰어 판단([3]-8)과 일치',
         '[3]-8로 통합', '종결([3]-8 통합)'],
        ['[3]-8\n(추가)', '리뷰어 추가 지적: 평가인은 내재가치 MAX(Ve+Vd−EXE, 0)와 시간가치(다음 단계 콜가치를 rd로 할인) 중 '
                        '큰 값을 콜가치로 산정 — 차액정산형 콜에 맞는 방법. 본 콜은 총액교환형(권리자가 행사가 총액 지급, '
                        '의무자가 CB 인도)이므로 W/WO가 타당. 수정 요청: ① T&F W/WO ② MIN[Call, MAX(만기상환, Put, Con, HV)] '
                        '③ 콜 행사기간 중 전환권 0 ④ Call로 산정되는 노드는 rf 할인',
         '(회신 대기)', '수정 요청',
         f'리뷰어 요청 방식 재계산(콜 대상 35%, 매매대금 지급기일 7회 버뮤다형, 2029-01-21은 Put과 동시 — MIN으로 콜 우선):\n'
         f'· 콜 없는 CB(Vwo) {base["V_wo"]:,.0f}원\n'
         f'· 매도청구권 = 35% × (Vwo − Vw) = {call_rf:,.0f}원 — 보고서 대비 {d_call:+,.0f}원({d_call / REP["매도청구권"]:+.1%})\n'
         f'· 전환제한(③) 미반영 시 {base["콜_WW"][("무위험이자율", False)]:,.0f}원 → 조건 ③의 영향이 큼\n'
         f'· Call 노드 rd 할인 시 {base["콜_WW"][("위험이자율", True)]:,.0f}원(Compound TF분할 '
         f'{base["콜_Compound"]["TF분할"]:,.0f}원과 같음 — 구현 검증)\n'
         '감사팀 의견: 총액교환형이라는 리뷰어 판단과 수정 방법에 동의(행사가 현금은 콜 권리자가 지급하므로 발행회사 신용위험 비반영)',
         '평가인 재평가 요청 → 수정 금액을 감사팀 재계산과 대사, 최초인식 분개(매도청구권 자산·전환권대가) 수정',
         '미결'],
    ])

    # ── 3. 재계산결과 ──
    ws = wb.create_sheet('재계산결과')
    s = Sheet(ws, [40, 20, 20, 18, 12, 50])
    s.title('리뷰 반영 재계산 결과 — 평가기준일 2026-07-21 (변동성 106.15%, 주간 52단계, 정상 보간, 콜 W/WO·rf)')
    rows = [
        ['일반사채(주계약)', REP['일반사채'], base['일반사채'], '[2]-3 보간 수정'],
        ['조기상환청구권(Put)', REP['조기상환청구권'], base['조기상환청구권'], '[2]-3 보간 수정'],
        ['전환권', REP['전환권'], base['전환권'], '변동성 106.15%(250일) — 종결'],
        ['전환사채(콜 제외)', REP['일반사채'] + REP['조기상환청구권'] + REP['전환권'], base['전환사채_콜제외'], '합계'],
        ['매도청구권(콜옵션, 회사 보유 자산)', REP['매도청구권'], call_rf, '[3]-8 W/WO·전환제한·rf'],
        ['전환사채 총가치(투자자 관점)', REP['총가치'], base['총가치'], '콜 제외 CB − 매도청구권'],
    ]
    st = s.table(['구성요소', '평가보고서(원)', '감사팀 재계산(원)', '차이(원)', '차이율', '비고'],
                 [[a, b, c, None, None, d] for a, b, c, d in rows], [None, WON, WON, WON, PCT])
    for i in range(len(rows)):
        r = st + i
        ws.cell(r, 4, f'=C{r}-B{r}').number_format = WON
        ws.cell(r, 5, f'=IF(B{r}=0,0,D{r}/B{r})').number_format = PCT
    s.blank()
    s.sub('최초인식 부채·자본 분리 영향 (회사: 조기상환권 파생상품부채 분리, 매도청구권 자산 인식 — 잔여법)')
    st = s.r + 1
    s.table(['항목', '평가보고서 기준(원)', '리뷰 반영(원)', '차이(원)', '', '산식'], [
        ['발행가액', ISSUE_PRICE, ISSUE_PRICE, None, None, '거래비용 0'],
        ['주계약(상각후원가 부채)', REP['일반사채'], base['일반사채'], None, None, ''],
        ['조기상환권(파생상품부채)', REP['조기상환청구권'], base['조기상환청구권'], None, None, ''],
        ['매도청구권(파생상품자산)', REP['매도청구권'], call_rf, None, None, ''],
        ['전환권대가(자본)', None, None, None, None, '발행가액 − 주계약 − 조기상환권 + 매도청구권'],
        ['전환권조정', None, None, None, None, '액면 − 주계약'],
    ], [None, WON, WON, WON])
    for r in range(st, st + 6):
        ws.cell(r, 4, f'=C{r}-B{r}').number_format = WON
    for col in ('B', 'C'):
        ws[f'{col}{st + 4}'] = f'={col}{st}-{col}{st + 1}-{col}{st + 2}+{col}{st + 3}'
        ws[f'{col}{st + 5}'] = f'={FACE}-{col}{st + 1}'
        ws[f'{col}{st + 4}'].number_format = ws[f'{col}{st + 5}'].number_format = WON
    s.note('※ 보고서 기준 전환권대가 27,790,115,375원·전환권조정 33,321,943,372원은 회사 장부와 일치(기존 검토). '
           '리뷰 반영 금액은 평가인 수정 보고서 수령 후 확정')

    # ── 4. 민감도분석 ──
    ws = wb.create_sheet('민감도분석')
    s = Sheet(ws, [52, 20, 20, 20, 14, 40])
    s.title('재계산 민감도 — 주간 격자(연 52단계), 보고서 p.25 KIS B-·국고채 매트릭스')
    s.sub('① [2]-3 YTM 보간 — 정상 보간 vs 만기 구간 한 칸 밀림(리뷰어 지적 유형) 재현')
    st = s.table(['구성요소', '평가보고서', '정상 보간(재계산)', '구간 밀림 재현', '밀림 영향', '비고'], [
        [k, REP[k], base[k], shifted[k], None, n] for k, n in
        (('일반사채', '과대할인 → 과소평가'), ('조기상환청구권', ''), ('전환권', ''))
    ], [None, WON, WON, WON, PCT])
    for r in range(st, st + 3):
        ws.cell(r, 5, f'=D{r}/C{r}-1').number_format = PCT
    s.note('※ 평가인 파일의 실제 매트릭스(3M 12.438%, 6M 13.914%)는 보고서 p.25와 달라 보고서 금액 자체는 재현되지 않음 — '
           '오류 유형의 금액 규모 판단용')
    s.blank()
    s.sub('② [3] 매도청구권 평가방법·할인방식 (변동성 106.15%, 콜 대상 35%)')
    ww, cp = base['콜_WW'], base['콜_Compound']
    st = s.table(['방식', '매도청구권(원)', '보고서 대비(원)', '보고서 대비(%)', '', '비고'], [
        ['W/WO + 전환제한 + Call 노드 rf 할인 (리뷰어 요청)', ww[('무위험이자율', True)], None, None, None, '재계산결과 시트 적용'],
        ['W/WO + 전환제한 + Call 노드 rd 할인', ww[('위험이자율', True)], None, None, None, 'Compound TF분할과 같은 값'],
        ['W/WO + 전환제한 없음 + rf 할인', ww[('무위험이자율', False)], None, None, None, '조건 ③ 미반영'],
        ['W/WO + 전환제한 없음 + rd 할인', ww[('위험이자율', False)], None, None, None, ''],
        ['Compound(별도격자) + 콜가치 전체 RD 할인', cp['위험이자율'], None, None, None, '평가인 방식 재현'],
        ['Compound + 콜가치 전체 RF 할인', cp['무위험이자율'], None, None, None, ''],
        ['Compound + TF분할(주식부분 RF, 현금부분 RD)', cp['TF분할'], None, None, None, ''],
        ['평가보고서', REP['매도청구권'], None, None, None, ''],
    ], [None, WON, WON, PCT])
    for r in range(st, st + 8):
        ws.cell(r, 3, f'=B{r}-{REP["매도청구권"]}').number_format = WON
        ws.cell(r, 4, f'=B{r}/{REP["매도청구권"]}-1').number_format = PCT
    s.note(f'※ 콜 없는 CB에 전환제한만 반영하면 {base["V_wo_전환제한"] - base["V_wo"]:+,.0f}원(무배당이라 조기전환 유인 없음) — '
           '전환제한 효과는 콜 행사 시 투자자가 전환으로 피할 수 없게 되는 데서 생김')
    s.blank()
    s.sub('③ [1] 변동성 참고 (종결 항목)')
    st = s.table(['변동성', '전환권', '보고서 대비', '매도청구권(W/WO·rf)', '총가치', '비고'], [
        ['106.15% (평가인 회신: 250일 연환산)', base['전환권'], base['전환권'] / REP['전환권'] - 1, base['콜'], base['총가치'], '적용'],
        ['105.73% (=106.15%×√(250/252))', v252['전환권'], v252['전환권'] / REP['전환권'] - 1, v252['콜'], v252['총가치'], '참고'],
    ], [None, WON, PCT3, WON, WON])

    # ── 5. 추가확인사항 ──
    s = Sheet(wb.create_sheet('추가확인사항'), [6, 12, 70, 12, 44, 20])
    s.title('평가인 요청자료·수정 요청 및 보고서 기재 오류')
    s.table(['No', '구분', '내용', '관련 코멘트', '회신/결과', '상태'], [
        [1, '요청자료', '주가변동성 계산 파일(180영업일 원자료, 연환산 일수)', '[1]', '평가상세파일 4번 4번 sheet, 250일', '종결'],
        [2, '요청자료', '신용등급 B- 출처 Screen Copy', '[2]-1', 'NICE평가정보(2024년말), 4번 1번 sheet', '종결'],
        [3, '요청자료', '부트스트래핑·노드별 YTM 계산 파일', '[2]-2', 'KIS-NET, 2·3번 파일(공모 배율 적용)', '종결'],
        [4, '수정 요청', 'YTM 보간 오류 수정(위험·무위험 모두) 후 주계약·조기상환권 재평가', '[2]-3', '회신 대기', '미결'],
        [5, '수정 요청', '매도청구권 W/WO(총액교환형)로 재평가 — MIN[Call, MAX(만기상환, Put, Con, HV)], 전환제한, Call rf 할인',
         '[3]-8', f'감사팀 재계산 {call_rf:,.0f}원', '미결'],
        [6, '기재 오류', 'p.26 조기상환권 행사가능기간 "발행일로부터 12개월(2027.07.21)" — 계약서 제4조 11호·p.7 표는 30개월(2029-01-21)',
         '-', '수정 보고서에서 정정 요청', '미결'],
        [7, '기재 오류', 'p.27 "유사기업의 주가 변동성" 표제 — 본문은 평가대상회사 주가 기준', '[1]', '수정 보고서에서 정정 요청', '미결'],
        [8, '기재 확인', 'p.25 매트릭스(3M 13.242%)와 평가인 파일 적용값(3M 12.438%) 차이 — 공모 배율 반영 매트릭스를 보고서에 기재 요청',
         '[2]-2·3', '수정 보고서에서 확인', '미결'],
        [9, '감사팀 판단', f'수정 보고서 수령 후 매도청구권 자산·전환권대가 수정 분개 검토(리뷰 반영 시 전환권대가 {eq_new - eq_rep:+,.0f}원)',
         '[3]-8', '', '미결'],
    ])

    for w in wb.worksheets:
        w.sheet_view.showGridLines = False
    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    wb.save(OUT)
    print(f'저장: {OUT}')
    print(f'주계약 {base["일반사채"]:,.0f} / 조기상환권 {base["조기상환청구권"]:,.0f} / 전환권 {base["전환권"]:,.0f}')
    print(f'매도청구권(W/WO·rf) {call_rf:,.0f} (보고서 대비 {d_call:+,.0f}) / 총가치 {base["총가치"]:,.0f}')
    print(f'전환권대가 {eq_rep:,.0f} → {eq_new:,.0f} ({eq_new - eq_rep:+,.0f})')


if __name__ == '__main__':
    build()
