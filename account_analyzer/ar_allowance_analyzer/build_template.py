"""표준 입력 템플릿(ar_allowance_template.xlsx) 생성 스크립트.

실행: python build_template.py
input_data/ar_allowance_template.xlsx 를 새로 만든다(이미 있으면 덮어씀).
회사별 파일은 이 템플릿을 복사해 ar_allowance_<company>_information_fy<year>.xlsx 로 저장해서 사용한다.

설계 원칙: 매출채권 연령분석 및 대손충당금 설정 검토.
  - 상장사(roll rate법)와 비상장사(연령별 대손율 설정법) 모두 "연령구간별 채권잔액에 구간별 대손율을
    곱해 대손충당금을 산출한다"는 계산 구조 자체는 동일하다 — 차이는 그 대손율을 어떻게 구했는지
    (상장사는 과거 이동(roll) 매트릭스로 산출, 비상장사는 실무관행상 연령별로 직접 설정)일 뿐이므로,
    엔진은 하나로 두고 '기준정보' 시트의 '상장구분'만 분기해 안내문구를 다르게 보여준다. roll rate
    자체(과거 여러 시점의 연령 이동 매트릭스로부터 대손율을 추정하는 과정)는 K-IFRS 계리보고서 앱과
    동일한 이유로 이 앱이 재현하지 않는다 — 통계적 모델을 근사 복제하면 차이가 오류인지 모델 단순화
    때문인지 구분할 수 없어 위험하다. 대신 회사(또는 계리/컨설팅 결과)가 이미 산출한 연령별 대손율을
    '연령별대손율' 시트에 입력받아 (a) 그 적용 산식(연령구간별 채권잔액×대손율=대손충당금)이 맞는지
    재계산으로 검증하고 (b) 연령구간이 커질수록 대손율이 감소하는 등 비정상 패턴이 있는지, 비상장사는
    최근 실제대손율과 괴리가 큰지를 확인하는 데 스코프를 한정한다.
  - 개별평가(부도/회생절차/소송 등 손상 징후가 있는 거래처)는 연령대와 무관하게 연령분석(집합평가)
    모집단에서 반드시 제외해야 한다 — 실무에서 가장 흔한 오류가 이 분리 누락이므로 '매출채권명세'
    시트에 개별평가대상여부 플래그를 두어 강제로 구분한다.
  - 특수관계자채권은 신용위험 성격이 일반 매출채권과 달라 집합평가 모집단에서 제외하고 별도 표시한다.
"""
import os
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_PATH = os.path.join(HERE, "input_data", "ar_allowance_template.xlsx")

YN_OPTIONS = ["Y", "N"]
LISTED_OPTIONS = ["상장", "비상장"]

DEFAULT_BUCKET_THRESHOLDS = "30,60,90,180,365"
DEFAULT_BUCKET_LABELS = ["정상(미도래)", "1~30일", "31~60일", "61~90일", "91~180일", "181~365일", "365일 초과"]

RATE_TABLE_MEMO = (
    "연령구간별 대손율 입력 메모\n\n"
    "'연령구간' 열은 '기준정보' 시트의 '연령구간 상한(일, 콤마구분)' 설정과 반드시 일치해야 합니다.\n"
    "기본값(30,60,90,180,365) 기준 구간은 왼쪽 예시행과 같이 7개입니다. 구간 상한을 바꾸면 이 표의\n"
    "행도 그에 맞춰 다시 작성해야 하며, 앱 실행 시 구간명이 일치하지 않으면 경고가 표시됩니다.\n\n"
    "'회사설정 대손율(%)'은 상장사는 roll rate법으로 산출한 결과값을, 비상장사는 회사가 실무관행상\n"
    "직접 설정한 연령별 대손율을 그대로 입력하면 됩니다 — 이 앱은 그 산출 과정(이동매트릭스 계산 등)\n"
    "자체를 재현하지 않고, 입력된 대손율이 채권잔액에 올바르게 곱해졌는지와 그 값이 합리적인지만 검증합니다.\n\n"
    "'최근 실제대손율(참고, 선택, %)'은 비상장사에서 특히 유용합니다 — 최근 3~5개년 정도의 그 연령\n"
    "구간에서 실제로 발생한 대손 실적을 대손율로 환산해 입력하면, 회사설정율과 자동 비교(back-test)되어\n"
    "과소/과대 설정 여부를 확인할 수 있습니다. 모르면 비워두세요(해당 비교만 생략됩니다)."
)

# ── '매출채권명세' 시트 ────────────────────────────────────────────────────
COLUMNS = [
    ("거래처정보", "거래처명", 18),
    ("거래처정보", "거래처코드(선택)", 14),
    ("거래처정보", "특수관계자여부(Y/N)", 14),
    ("채권현황", "채권잔액(원)", 16),
    ("채권현황", "연령산정 기산일(결제기일/만기일)", 20),
    ("채권현황", "회사계산 경과일수(선택,참고용)", 18),
    ("개별평가(부도/회생/소송 등)", "개별평가대상여부(Y/N)", 16),
    ("개별평가(부도/회생/소송 등)", "개별평가사유(선택)", 20),
    ("개별평가(부도/회생/소송 등)", "개별평가 회수가능예상액(선택,원)", 20),
    ("신용보강(선택)", "담보/보증 등 차감액(원)", 16),
    ("대사용(선택)", "거래처별 회사계상 대손충당금(원)", 20),
    ("기타", "비고", 30),
]

MONEY_FIELDS = ("채권잔액(원)", "개별평가 회수가능예상액(선택,원)", "담보/보증 등 차감액(원)",
                "거래처별 회사계상 대손충당금(원)")
DAY_FIELDS = ("회사계산 경과일수(선택,참고용)",)
DATE_FIELDS = ("연령산정 기산일(결제기일/만기일)",)

EXAMPLES = [
    ["㈜예시)정상거래처", "CUST-001", "N",
     50000000, "2026-11-15", None,
     "N", None, None,
     None, None,
     "예시) 결산기준일이 2026-12-31이면 경과일수 46일 → '31~60일' 구간에 자동 분류"],
    ["㈜예시)관계사", "CUST-002", "Y",
     30000000, "2026-06-30", None,
     "N", None, None,
     None, None,
     "예시) 특수관계자채권 — 집합평가(연령분석) 모집단에서 제외되고 별도 표로 표시됨"],
    ["㈜예시)회생절차거래처", "CUST-003", "N",
     80000000, "2025-03-01", None,
     "Y", "회생절차 개시(2026-05월)", 20000000,
     None, 60000000,
     "예시) 개별평가대상 — 연령분석 대신 채권잔액-회수가능예상액=대손충당금(60,000,000원)으로 개별 계산됨"],
    ["㈜예시)장기연체거래처", "CUST-004", "N",
     12000000, "2025-10-01", 456,
     "N", None, None,
     2000000, None,
     "예시) 담보차감 후 순채권액(10,000,000원)에 '365일 초과' 구간 대손율 적용. 회사계산 경과일수를"
     " 입력하면 앱 계산 경과일수와 자동 대사됨"],
]


def build():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "매출채권명세"

    group_fill = PatternFill("solid", fgColor="D9E1F2")
    header_fill = PatternFill("solid", fgColor="4472C4")
    header_font = Font(bold=True, color="FFFFFF")
    group_font = Font(bold=True)
    thin = Side(style="thin", color="B7B7B7")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    input_fill = PatternFill("solid", fgColor="FFF2CC")

    def _col_letter(cols, header_name: str) -> str:
        idx = next(i for i, (_, name, _) in enumerate(cols, start=1) if name == header_name)
        return get_column_letter(idx)

    # 1행: 그룹 헤더(병합), 2행: 상세 헤더
    start = 1
    prev_group = COLUMNS[0][0]
    for i, (group, _, _) in enumerate(COLUMNS, start=1):
        if group != prev_group:
            ws.merge_cells(start_row=1, start_column=start, end_row=1, end_column=i - 1)
            start = i
            prev_group = group
    ws.merge_cells(start_row=1, start_column=start, end_row=1, end_column=len(COLUMNS))

    for i, (group, name, width) in enumerate(COLUMNS, start=1):
        c1 = ws.cell(row=1, column=i, value=group if i == 1 or COLUMNS[i - 2][0] != group else None)
        c1.fill = group_fill
        c1.font = group_font
        c1.alignment = center
        c1.border = border

        c2 = ws.cell(row=2, column=i, value=name)
        c2.fill = header_fill
        c2.font = header_font
        c2.alignment = center
        c2.border = border
        ws.column_dimensions[get_column_letter(i)].width = width

    ws.row_dimensions[1].height = 20
    ws.row_dimensions[2].height = 32
    ws.freeze_panes = "A3"

    for r, row in enumerate(EXAMPLES, start=3):
        for c, val in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=val)
            cell.border = border
            field = COLUMNS[c - 1][1]
            if field in DATE_FIELDS and val:
                cell.number_format = "yyyy-mm-dd"
            if field in MONEY_FIELDS and val is not None:
                cell.number_format = "#,##0"
            if field in DAY_FIELDS and val is not None:
                cell.number_format = "0"

    last_row = 1000
    for label, options, error in [
        ("특수관계자여부(Y/N)", YN_OPTIONS, "Y 또는 N 중 선택하세요."),
        ("개별평가대상여부(Y/N)", YN_OPTIONS, "Y 또는 N 중 선택하세요."),
    ]:
        dv = DataValidation(type="list", formula1=f'"{",".join(options)}"', allow_blank=True, showErrorMessage=True)
        dv.error = error
        ws.add_data_validation(dv)
        col = _col_letter(COLUMNS, label)
        dv.add(f"{col}3:{col}{last_row}")

    # ── '연령별대손율' 시트 ──────────────────────────────────────────────
    RATE_COLUMNS = [("연령구간", 14), ("회사설정 대손율(%)", 16),
                     ("최근 실제대손율(참고,선택,%)", 18), ("비고", 30)]
    ws_rate = wb.create_sheet("연령별대손율")
    for i, (name, width) in enumerate(RATE_COLUMNS, start=1):
        c = ws_rate.cell(row=1, column=i, value=name)
        c.fill = header_fill
        c.font = header_font
        c.alignment = center
        c.border = border
        ws_rate.column_dimensions[get_column_letter(i)].width = width
        if name == "회사설정 대손율(%)":
            memo = Comment(RATE_TABLE_MEMO, "ar_allowance_analyzer")
            memo.width, memo.height = 420, 320
            c.comment = memo
    ws_rate.row_dimensions[1].height = 32
    ws_rate.freeze_panes = "A2"

    RATE_EXAMPLES = [
        ["정상(미도래)", 0.5, 0.3, "예시) 아직 결제기일이 도래하지 않은 채권도 K-IFRS9 기대신용손실모형·"
                                    "일반기업회계기준 실무관행상 소액이나마 대손율을 설정하는 것이 일반적"],
        ["1~30일", 1.0, 0.8, None],
        ["31~60일", 3.0, 2.5, None],
        ["61~90일", 8.0, 6.0, None],
        ["91~180일", 20.0, 15.0, None],
        ["181~365일", 50.0, 40.0, None],
        ["365일 초과", 100.0, 90.0, "예시) 통상 1년 초과 연체분은 100% 설정하는 경우가 많음(회사 정책에 따라 다름)"],
    ]
    for r, row in enumerate(RATE_EXAMPLES, start=2):
        for c, val in enumerate(row, start=1):
            cell = ws_rate.cell(row=r, column=c, value=val)
            cell.border = border
            if c in (2, 3) and val is not None:
                cell.number_format = "0.0"
                cell.fill = input_fill

    # ── '기준정보' 시트 ──────────────────────────────────────────────────
    basis = wb.create_sheet("기준정보")
    basis.column_dimensions["A"].width = 42
    basis.column_dimensions["B"].width = 20
    basis.column_dimensions["C"].width = 55

    basis.cell(row=1, column=1, value="매출채권 대손충당금(연령분석) 기준정보").font = Font(bold=True, size=13)

    header_row = 3
    for i, h in enumerate(["항목", "값", "설명"], start=1):
        cell = basis.cell(row=header_row, column=i, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border

    rows = [
        ("상장구분(상장/비상장)", "비상장",
         "상장사는 통상 roll rate법(과거 연령 이동 매트릭스)으로, 비상장사는 연령별 대손율을 직접 설정하는 "
         "방식으로 대손충당금을 산정합니다. 이 앱의 계산 구조(연령구간별 채권잔액×대손율)는 두 경우 모두 "
         "동일하며, 이 값은 안내문구·강조 항목만 다르게 표시하는 용도입니다(예: 비상장사는 '최근 실제대손율' "
         "대비 검증을, 상장사는 아래 'Forward-looking 조정 반영 여부'를 강조).", "listbox"),
        ("연령구간 상한(일, 콤마구분)", DEFAULT_BUCKET_THRESHOLDS,
         "예: 30,60,90,180,365 → 정상(미도래)/1~30일/31~60일/61~90일/91~180일/181~365일/365일 초과 7개 구간이 "
         "생성됩니다. 바꾸면 '연령별대손율' 시트의 구간명도 반드시 맞춰 다시 작성해야 합니다.", "text"),
        ("Forward-looking(미래전망정보) 조정 반영 여부(Y/N, 상장사 참고)", "N",
         "K-IFRS9 기대신용손실모형은 과거 실적(roll rate) 기반 대손율에 거시경제지표 등 미래전망정보를 "
         "가산 조정하도록 요구합니다. 이 앱은 그 조정치를 재계산하지 않으므로, 회사가 반영했는지 여부와 "
         "근거 문서화 상태만 이 항목에 표시해 감사조서에 남깁니다.", "listbox"),
        ("전기말 회사계상 매출채권 총액(원)", None,
         "전기말 재무제표상 매출채권 총액. 아래 전기말 대손충당금과 함께 전기 설정률을 구해 당기 설정률과 "
         "자동 비교합니다(설정률 급변 시 감사절차상 사유 확인 필요).", "money"),
        ("전기말 회사계상 대손충당금(원)", None,
         "전기말 재무제표(또는 결산정산표)상 실제 계상액.", "money"),
        ("당기말 회사계상 대손충당금(원)", None,
         "당기말 재무제표상 회사 계상액. 앱의 재계산 합계와 비교해 차이(대사)를 표시합니다.", "money"),
        ("당기 대손충당금 전입액(손익, 선택, 원)", None,
         "당기 손익계산서상 대손상각비 중 충당금 전입액. 아래 T계정 검증(tie-out)에 사용됩니다.", "money"),
        ("당기 대손충당금 환입액(선택, 원)", None,
         "당기 중 대손충당금 환입액.", "money"),
        ("당기 대손금 직접상각(제각)액(선택, 분개장 기준, 원)", None,
         "journal_analyzer(분개장분석)에서 대손충당금(또는 매출채권) 계정의 당기 직접상각(제각) 분개 합계를 "
         "뽑아 입력. 전기말(회사계상)+당기전입액-당기환입액-이 값이 당기말(회사계상)과 맞는지 요약표에서 "
         "자동 검증(T계정 tie-out)합니다. 모르면 비워둬도 됩니다(해당 검증만 생략).", "money"),
    ]

    r = header_row + 1
    for label, default, note, kind in rows:
        c1 = basis.cell(row=r, column=1, value=label)
        c1.border = border
        c2 = basis.cell(row=r, column=2, value=default if kind != "money" else None)
        c2.border = border
        c2.alignment = Alignment(horizontal="center", vertical="center")
        if kind == "money":
            c2.number_format = "#,##0"
            c2.fill = input_fill
        elif kind == "text":
            c2.fill = input_fill
        elif kind == "listbox":
            c2.fill = input_fill
            options = LISTED_OPTIONS if "상장구분" in label else YN_OPTIONS
            dv = DataValidation(type="list", formula1=f'"{",".join(options)}"', allow_blank=True, showErrorMessage=True)
            dv.error = f"{'/'.join(options)} 중 선택하세요."
            basis.add_data_validation(dv)
            dv.add(f"B{r}")
        c3 = basis.cell(row=r, column=3, value=note)
        c3.border = border
        c3.alignment = Alignment(wrap_text=True, vertical="top")
        r += 1

    basis.cell(row=r + 1, column=1,
               value="※ 결산기준일(당기말/전기말)은 이 시트에 입력하지 않습니다 — 실행 시 파일명(fy<연도>)과 "
                     "--fiscal-month 옵션으로 다른 계정 검증앱과 동일한 규칙으로 자동 결정됩니다.").font = \
        Font(italic=True, color="808080")

    # ── 안내 시트 ────────────────────────────────────────────────────────
    guide = wb.create_sheet("작성안내")
    guide.column_dimensions["A"].width = 100
    lines = [
        "매출채권 연령분석 및 대손충당금 설정 검증앱 — 입력 템플릿 작성 안내",
        "",
        "핵심 계산식",
        "  경과일수 = 결산기준일 − 연령산정 기산일(결제기일/만기일)",
        "  순채권액 = 채권잔액 − 담보/보증 등 차감액",
        "  [집합평가 대상] 대손충당금(계산) = 순채권액 × 해당 연령구간의 대손율(연령별대손율 시트)",
        "  [개별평가 대상] 대손충당금(계산) = 순채권액 − 개별평가 회수가능예상액(미입력 시 순채권액 전액을 "
        "잠정 대손충당금으로 계상하고 경고 표시)",
        "  [특수관계자채권] 신용위험 성격이 달라 위 계산에서 제외하고 별도 표로만 표시(대손충당금 계산은 별도 검토)",
        "",
        "1. '매출채권명세' 시트 — 거래처(또는 채권 건) 1개 = 1행.",
        "   '특수관계자여부': Y로 표시하면 연령분석(집합평가) 모집단에서 제외되고 별도 표에 표시됩니다.",
        "   '개별평가대상여부': 부도·회생절차·소송 등 손상 징후가 있는 거래처는 반드시 Y로 표시하세요.",
        "     연령대와 무관하게 연령분석에서 제외되고, '개별평가 회수가능예상액'과의 차액으로 대손충당금이",
        "     별도 계산됩니다. 회수가능예상액을 모르면 비워두되, 이 경우 순채권액 전액이 잠정 대손충당금으로",
        "     계상되고 요약표에 경고가 표시되니 실제 회수가능액을 파악해 채워 넣는 것을 권장합니다.",
        "   '연령산정 기산일': 세금계산서 발행일이 아니라 '결제기일(만기일)'을 기준으로 경과일수를 계산합니다",
        "     (여신기간을 이미 반영한 날짜). 회사 매출채권 연령분석표에 보통 이미 이 날짜 기준 경과일수가",
        "     있으므로, 그 경과일수를 '회사계산 경과일수(선택)'에 함께 입력하면 앱 계산과 자동 대사됩니다.",
        "   '담보/보증 등 차감액': 담보나 지급보증 등으로 신용위험이 상쇄되는 금액이 있으면 입력하세요",
        "     (순채권액에서 차감된 후 대손율이 곱해집니다). 없으면 비워두세요.",
        "   '거래처별 회사계상 대손충당금': 회사가 거래처별로 이미 대손충당금을 계산해뒀다면 입력하세요 —",
        "     앱 계산값과 거래처별로 자동 대사되어 차이가 표시됩니다. 총액만 아는 경우 비워두고 '기준정보'",
        "     시트의 '당기말 회사계상 대손충당금(원)'에 총액만 입력해도 총액 대사는 됩니다.",
        "",
        "2. '연령별대손율' 시트 — '기준정보'의 '연령구간 상한' 설정과 일치하는 구간명으로 행을 구성하세요",
        "   (기본값 기준 7개 구간 예시가 이미 채워져 있습니다). '회사설정 대손율(%)'은 상장사는 roll rate법",
        "   결과값을, 비상장사는 회사가 실무관행상 설정한 연령별 대손율을 그대로 입력하면 됩니다 — 이 앱은",
        "   그 산출 과정 자체를 재현하지 않고, 채권잔액에 올바르게 적용됐는지와 그 값이 합리적인지만",
        "   검증합니다(연령구간이 커질수록 대손율이 감소하는 등 비정상 패턴은 자동 경고). '최근 실제대손율",
        "   (참고, 선택)'을 입력하면 회사설정율과 자동 비교(back-test)됩니다 — 특히 비상장사에서 과소/과대",
        "   설정 여부를 확인하는 데 유용합니다.",
        "",
        "3. '기준정보' 시트",
        "   '상장구분': 안내문구만 다르게 표시할 뿐 계산 로직 자체는 상장/비상장 동일합니다.",
        "   '연령구간 상한(일, 콤마구분)': 바꾸면 '연령별대손율' 시트도 반드시 맞춰 다시 작성하세요.",
        "   'Forward-looking 조정 반영 여부': 상장사(K-IFRS9)는 과거실적 기반 대손율에 미래전망정보를",
        "     가산 조정해야 합니다. 이 앱은 그 조정치를 재계산하지 않고, 회사가 반영했는지 여부만 감사조서에",
        "     기록하는 용도입니다.",
        "   전기말/당기말 회사계상 금액들과 당기 전입액/환입액/직접상각액을 입력하면 설정률 전기 대비 비교와",
        "   T계정 검증(tie-out)이 요약표에 자동으로 표시됩니다. 모르는 항목은 비워두면 해당 검증만 생략됩니다.",
        "",
        "4. 결산기준일은 이 파일에 입력하지 않습니다. 실행 시 파일명의 'fy<연도>'와 --fiscal-month 옵션(기본",
        "   12월)으로 depreciation_analyzer 등 다른 계정 검증앱과 동일한 규칙으로 자동 계산합니다.",
        "",
        "5. 파일명 규칙: 이 템플릿을 복사해 'ar_allowance_<회사명>_information_fy<회계연도>.xlsx' 로 저장하세요.",
        "   예) ar_allowance_kyungnam_information_fy2026.xlsx",
    ]
    for i, line in enumerate(lines, start=1):
        cell = guide.cell(row=i, column=1, value=line)
        if i == 1:
            cell.font = Font(bold=True, size=13)
        cell.alignment = Alignment(wrap_text=True, vertical="top")

    os.makedirs(os.path.join(HERE, "input_data"), exist_ok=True)
    wb.save(OUT_PATH)
    print(f"템플릿 생성 완료: {OUT_PATH}")


if __name__ == "__main__":
    build()
