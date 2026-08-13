"""표준 입력 템플릿(depreciation_template.xlsx) 생성 스크립트.

실행: python build_template.py
input_data/depreciation_template.xlsx 를 새로 만든다(이미 있으면 덮어씀).
회사별 파일은 이 템플릿을 복사해 depreciation_<company>_information_fy<year>.xlsx 로 저장해서 사용한다.
"""
import os
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
OUT_PATH = os.path.join(HERE, "input_data", "depreciation_template.xlsx")

# (그룹헤더, 상세헤더, 열너비)
COLUMNS = [
    ("자산기본정보", "자산관리번호", 12),
    ("자산기본정보", "계정과목", 16),
    ("자산기본정보", "자산명(세부내역)", 28),
    ("취득정보", "취득일", 12),
    ("취득정보", "취득원가", 14),
    ("취득정보", "잔존가치", 12),
    ("취득정보", "내용연수(년)", 10),
    ("상각정보", "상각방법(정액법/정률법)", 12),
    ("상각정보", "상각률(정률법전용)", 12),
    ("상각정보", "상각개시(당월/익월)", 12),
    ("처분정보", "처분일", 12),
    ("손상차손(선택)", "손상차손 인식일", 12),
    ("손상차손(선택)", "손상차손 인식액", 14),
    ("회사계상액(대사용)", "전기말 회사계상 감가상각누계액", 16),
    ("회사계상액(대사용)", "당기 회사계상 감가상각비", 14),
    ("기타", "원가구분", 10),
    ("기타", "비고", 20),
]

ACCOUNT_SUGGESTIONS = [
    "건물", "건물(투자)", "구축물", "기계장치", "차량운반구",
    "비품", "시설장치", "사용권자산", "투자부동산",
]


def build():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "자산정보"

    group_fill = PatternFill("solid", fgColor="D9E1F2")
    header_fill = PatternFill("solid", fgColor="4472C4")
    header_font = Font(bold=True, color="FFFFFF")
    group_font = Font(bold=True)
    thin = Side(style="thin", color="B7B7B7")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 1행: 그룹 헤더 (병합)
    col = 1
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

    # 예시행 (정액법 1건 + 정률법 1건 + 손상차손 반영 1건)
    examples = [
        [
            "FA-0001", "기계장치", "예시) 사출성형기 1호", "2023-03-15", 120000000, 0, 8,
            "정액법", None, "당월", None, None, None, 45000000, 15000000, "제조", "예시 행 — 실제 자산으로 교체",
        ],
        [
            "FA-0002", "차량운반구", "예시) 업무용 승용차", "2024-07-01", 45000000, 0, 5,
            "정률법", 0.451, "익월", None, None, None, 0, 10147500, "판관", "예시 행 — 실제 자산으로 교체",
        ],
        [
            "FA-0003", "건물", "예시) CGU 손상평가 반영 자산", "2015-01-01", 500000000, 0, 40,
            "정액법", None, "당월", None, "2025-12-31", 80000000, None, None, "제조",
            "예시) 25년말 CGU 손상평가로 손상차손 인식 → 남은 기간(내용연수 그대로) 정액법 재상각",
        ],
    ]
    for r, row in enumerate(examples, start=3):
        for c, val in enumerate(row, start=1):
            cell = ws.cell(row=r, column=c, value=val)
            cell.border = border
            if COLUMNS[c - 1][1] in ("취득일", "손상차손 인식일") and val:
                cell.number_format = "yyyy-mm-dd"
            if COLUMNS[c - 1][1] in ("취득원가", "잔존가치", "손상차손 인식액",
                                     "전기말 회사계상 감가상각누계액", "당기 회사계상 감가상각비"):
                cell.number_format = "#,##0"
            if COLUMNS[c - 1][1] == "상각률(정률법전용)" and val is not None:
                cell.number_format = "0.000"

    # 데이터 유효성 검사 (100행까지)
    last_row = 100

    dv_method = DataValidation(type="list", formula1='"정액법,정률법"', allow_blank=True, showErrorMessage=True)
    dv_method.error = "정액법 또는 정률법 중 선택하세요."
    ws.add_data_validation(dv_method)
    dv_method.add(f"H3:H{last_row}")

    dv_start = DataValidation(type="list", formula1='"당월,익월"', allow_blank=True, showErrorMessage=True)
    dv_start.error = "당월 또는 익월 중 선택하세요."
    ws.add_data_validation(dv_start)
    dv_start.add(f"J3:J{last_row}")

    dv_account = DataValidation(
        type="list",
        formula1=f'"{",".join(ACCOUNT_SUGGESTIONS)}"',
        allow_blank=True,
        showErrorMessage=False,  # 목록 외 자유 입력 허용 (신규 계정과목 대응)
    )
    ws.add_data_validation(dv_account)
    dv_account.add(f"B3:B{last_row}")

    # 안내 시트
    guide = wb.create_sheet("작성안내")
    guide.column_dimensions["A"].width = 100
    lines = [
        "감가상각비 검증앱 — 입력 템플릿 작성 안내",
        "",
        "1. 자산 1건 = 1행. '자산정보' 시트에 계속 추가하면 됩니다(계정과목별로 시트를 나누지 않음).",
        "2. 계정과목: 드롭다운 목록에 없는 계정(신규 취득 등)은 직접 입력해도 됩니다.",
        "3. 상각방법: 정액법 선택 시 '상각률' 열은 비워두면 자동 계산됩니다((취득원가-잔존가치)/내용연수).",
        "   정률법 선택 시 '상각률'을 반드시 입력하세요(세법상 상각률표 등 참고).",
        "4. 상각개시(당월/익월): 취득월부터 상각을 시작하면 '당월', 취득 다음달부터 시작하면 '익월'.",
        "   (리스 스케줄 앱과 동일한 정책 — 회사 관행에 맞게 계약/자산별로 다르게 지정 가능)",
        "5. 처분일: 당기 중 처분/폐기된 자산만 입력. 처분월까지만 상각하고 그 이후는 계산에서 제외됩니다.",
        "6. 손상차손(선택) — CGU 평가 등으로 손상차손을 인식한 자산만 입력(해당 없으면 비워둠).",
        "   '손상차손 인식일'이 속한 월까지는 원래 방식대로 정상 상각한 뒤, 그 시점의 장부금액에서 '손상차손 인식액'을 차감합니다.",
        "   이후 남은 기간은 '내용연수(년)'에서 정한 원래 종료시점까지 그대로 두고(잔존내용연수 재평가 없음), 정액법 자산은 (손상후 장부금액-잔존가치)/남은 개월수로 상각액을 다시 계산합니다.",
        "   정률법 자산은 손상 이후에도 같은 상각률을 손상후 장부금액에 계속 적용합니다.",
        "   감가상각누계액과 손상차손누계액은 출력물에서 별도 항목으로 표시됩니다(취득원가-감가상각누계액-손상차손누계액=장부금액).",
        "7. 회사계상액(대사용) — '당기 회사계상 감가상각비'를 채워두면, 앱이 재계산한 금액과 자동 비교해 차이를 표시합니다.",
        "   모르면 비워둬도 앱 실행에는 문제 없습니다(대사만 생략됨).",
        "8. 파일명 규칙: 이 템플릿을 복사해 'depreciation_<회사명>_information_fy<회계연도>.xlsx' 로 저장하세요.",
        "   예) depreciation_kyungnam_information_fy2026.xlsx",
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
