"""퇴직급여충당부채 검증앱(일반기업회계기준·퇴직금 추계액 방식) — 엔진.

input_data/severance_<company>_information_fy<year>.xlsx 를 읽어
결산기준일 현재 재직 중인 임직원별 퇴직금 추계액을 재계산하고,
전기말/당기말 회사계상 충당부채와 대사하는 output/severance_schedule_<company>_<fy>.xlsx 를 생성한다.

핵심 설계: depreciation_analyzer와 달리 기초(전기말) 잔액을 신뢰하지 않는다. 당기말 퇴충잔액은
결산기준일 현재 재직 중인 임직원의 재직일수·급여만으로 전액 독립 재계산하며, 회사계상 전기말/당기말
잔액은 오직 대사(비교)용 참고값으로만 쓰인다. 결산기준일은 입력파일에 직접 적지 않고
파일명(fy<연도>)과 --fiscal-month 로 depreciation_analyzer와 동일한 규칙으로 계산한다.

전기/당기 인원은 회사별로 별도 파일을 만드는 대신, 한 파일 안에 '전기정보'/'당기정보'
두 시트로 나눠 입력받는다(회사가 매년 스냅샷 형태로 제공하는 인원현황을 그대로 붙여넣는 것을 전제).
당기 계산(추계액)은 '당기정보' 시트만 사용하고, '전기정보'는 사번(없으면 성명) 기준으로
당기 시트와 매칭해 신규입사자/퇴사자 명단을 산출하는 데에만 쓰인다 — 실제 퇴사일 필드가 없어도
"전기에는 있었는데 당기에는 없는 사번"으로 진짜 퇴사자를 식별할 수 있다.

퇴직금 추계액 산식(근로기준법 간편 산식, '당기정보' 시트 기준):
  급여기준액 = 월평균급여 + 상여금(연간) × 3 / 12   (상여금 최근 3개월분 상당액을 월평균급여에 가산)
  1일평균임금 = 급여기준액 × 12 / 365
  퇴직금 추계액 = 1일평균임금 × 30일 × (재직일수 / 365) × 배수
  배수: 부서명이 정확히 '임원'인 인원에만 적용(임원퇴직금지급규정 등에 따른 배수, 일반 직원은 항상 1배)
  재직일수 = 당기 결산기준일 − 기산일
  기산일 = 중간정산일(퇴직금 중간정산 이력이 있는 경우) 있으면 그 날짜, 없으면 입사일
    ※ 중간정산은 재직 상태를 유지한 채 근속기간만 새로 기산되는 것이므로(퇴사가 아님),
      재직일수 계산 기산일에만 반영한다.

실행 예:
    python severance_schedule.py kyungnam --fiscal-month 12
    python severance_schedule.py --file severance_kyungnam_information_fy2026.xlsx
"""
import argparse
import calendar
import glob
import os
from datetime import date

import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(HERE, "input_data")
OUTPUT_DIR = os.path.join(HERE, "output")

SIG_THRESHOLD_ABS = 1000      # 유의차이 절대금액 기준(원)
SIG_THRESHOLD_PCT = 0.01      # 유의차이 비율 기준(1%)

COST_TYPES = ["제조원가", "판관비"]
BASIS_KEYS = {
    "제조원가": {
        "전기": "전기말 회사계상 퇴직급여충당부채(제조원가분)",
        "당기": "당기말 회사계상 퇴직급여충당부채(제조원가분)",
    },
    "판관비": {
        "전기": "전기말 회사계상 퇴직급여충당부채(판관비분)",
        "당기": "당기말 회사계상 퇴직급여충당부채(판관비분)",
    },
}

EXECUTIVE_DEPT = "임원"  # 부서명이 정확히 이 값일 때만 배수를 적용한다.

CURRENT_SHEET = "당기정보"
PRIOR_SHEET = "전기정보"

DISPLAY_COLS = ["사업장", "부서", "사번", "성명", "직급", "원가구분", "입사일", "중간정산일"]


# ── 공용 헬퍼 ────────────────────────────────────────────────────────────────

def _safe_float(v, default: float = 0.0) -> float:
    if v is None or v == "":
        return default
    try:
        if pd.isna(v):
            return default
    except (TypeError, ValueError):
        pass
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def _safe_date(v):
    if v is None or v == "":
        return None
    try:
        if pd.isna(v):
            return None
    except (TypeError, ValueError):
        pass
    ts = pd.Timestamp(v)
    return ts.date() if not pd.isna(ts) else None


def _cost_type(emp: dict) -> str:
    return str(emp.get("원가구분(제조원가/판관비)") or "").strip() or "(미분류)"


def _employee_key(emp: dict):
    """전기/당기 인원 매칭 키. 사번이 있으면 사번, 없으면 성명으로 매칭한다."""
    사번 = emp.get("사번")
    if 사번 not in (None, ""):
        return f"ID::{str(사번).strip()}"
    성명 = emp.get("성명")
    if 성명 not in (None, ""):
        return f"NAME::{str(성명).strip()}"
    return None


def _fy_bounds(target_fy: str, fiscal_month: int) -> tuple:
    """대상 회계연도(target_fy)의 시작월/종료월(YYYY-MM 문자열) 반환.
    depreciation_schedule.py/lease_schedule.py 와 동일 규칙."""
    fy = int(target_fy)
    if fiscal_month == 12:
        return f"{fy}-01", f"{fy}-12"
    return f"{fy - 1}-{fiscal_month + 1:02d}", f"{fy}-{fiscal_month:02d}"


def _apply_interim(fy_end: str, target_fy: str, interim_month: int = None) -> str:
    if not interim_month:
        return fy_end
    interim_ym = f"{target_fy}-{interim_month:02d}"
    return min(fy_end, interim_ym)


def _ym_to_end_date(ym: str) -> date:
    y, m = int(ym[:4]), int(ym[5:7])
    last_day = calendar.monthrange(y, m)[1]
    return date(y, m, last_day)


def _find_input_file(company: str = None, file: str = None) -> str:
    if file:
        path = file if os.path.isabs(file) else os.path.join(INPUT_DIR, file)
        if not os.path.exists(path):
            raise FileNotFoundError(f"입력 파일을 찾을 수 없습니다: {path}")
        return path

    if company:
        pattern = os.path.join(INPUT_DIR, f"severance_{company}_information_fy*.xlsx")
    else:
        pattern = os.path.join(INPUT_DIR, "severance_*_information_fy*.xlsx")

    matches = [p for p in glob.glob(pattern) if "template" not in os.path.basename(p)]
    if not matches:
        raise FileNotFoundError(f"입력 파일을 찾을 수 없습니다: {pattern}")
    if len(matches) > 1 and not company:
        raise ValueError(f"회사를 특정해주세요. 후보 파일 여러 개: {matches}")
    return matches[0]


def _is_significant(diff, base) -> bool:
    if diff is None or pd.isna(diff):
        return False
    if abs(diff) >= SIG_THRESHOLD_ABS and (base in (None, 0) or abs(diff) >= abs(base) * SIG_THRESHOLD_PCT):
        return True
    return False


# ── 입력 로딩 ────────────────────────────────────────────────────────────────

def load_employees(path: str, sheet_name: str) -> list:
    """지정한 인원 시트(1~2행 헤더, 3행부터 데이터)를 읽어 dict 목록으로 반환.
    시트가 없으면 빈 목록을 반환한다(예: '전기정보' 미작성 시 인원변동 명단만 생략됨)."""
    wb = openpyxl.load_workbook(path, data_only=True)
    if sheet_name not in wb.sheetnames:
        return []
    ws = wb[sheet_name]
    headers = [c.value for c in ws[2]]
    employees = []
    for row in ws.iter_rows(min_row=3, values_only=True):
        if row is None or all(v is None for v in row):
            continue
        rec = dict(zip(headers, row))
        if not rec.get("사번") and not rec.get("성명"):
            continue
        employees.append(rec)
    return employees


def load_basis(path: str) -> dict:
    """'기준정보' 시트를 라벨(A열) 기준으로 읽어 {라벨: 금액} dict 반환. 없으면 빈 dict."""
    wb = openpyxl.load_workbook(path, data_only=True)
    if "기준정보" not in wb.sheetnames:
        return {}
    ws = wb["기준정보"]
    basis = {}
    for row in ws.iter_rows(values_only=True):
        if not row or len(row) < 2:
            continue
        label, value = row[0], row[1]
        if not isinstance(label, str):
            continue
        label = label.strip()
        if not label:
            continue
        basis[label] = None if value in (None, "") else _safe_float(value)
    return basis


# ── 전기/당기 인원 매칭 (신규입사자/퇴사자 판정) ──────────────────────────────

def match_periods(당기_employees: list, 전기_employees: list) -> dict:
    """당기/전기 인원을 사번(없으면 성명) 기준으로 매칭해 신규입사자/퇴사자를 산출한다."""
    당기_by_key, 전기_by_key = {}, {}
    for e in 당기_employees:
        k = _employee_key(e)
        if k is not None:
            당기_by_key.setdefault(k, e)
    for e in 전기_employees:
        k = _employee_key(e)
        if k is not None:
            전기_by_key.setdefault(k, e)

    신규입사_keys = set(당기_by_key) - set(전기_by_key)
    퇴사_keys = set(전기_by_key) - set(당기_by_key)

    return {
        "신규입사_keys": 신규입사_keys,
        "신규입사자": [당기_by_key[k] for k in 신규입사_keys],
        "퇴사자": [전기_by_key[k] for k in 퇴사_keys],
        "전기인원수": len(전기_by_key),
        "당기인원수": len(당기_by_key),
    }


def _to_display_df(records: list) -> pd.DataFrame:
    rows = []
    for e in records:
        rows.append({
            "사업장": e.get("사업장") or "",
            "부서": e.get("부서") or "",
            "사번": e.get("사번"),
            "성명": e.get("성명"),
            "직급": e.get("직급"),
            "원가구분": _cost_type(e),
            "입사일": e.get("입사일"),
            "중간정산일": e.get("중간정산일"),
        })
    if not rows:
        return pd.DataFrame(columns=DISPLAY_COLS)
    return pd.DataFrame(rows)[DISPLAY_COLS]


# ── 인원별 퇴직금 추계액 계산 ('당기정보' 시트 기준) ──────────────────────

def compute_employee(emp: dict, 당기결산일: date) -> dict:
    입사일 = _safe_date(emp.get("입사일"))
    중간정산일 = _safe_date(emp.get("중간정산일"))
    월평균급여 = _safe_float(emp.get("월평균급여(원)"))
    연간상여금 = _safe_float(emp.get("상여금(연간, 원)"))
    부서 = str(emp.get("부서") or "").strip()
    # 상여금은 연 1회 등 정기 지급을 가정, 근로기준법상 평균임금 산정 관행대로
    # "최근 3개월분 상당액(연간상여금×3/12)"만 월평균급여에 가산한다.
    급여기준액 = 월평균급여 + 연간상여금 * 3 / 12

    result = {
        "재직일수": None, "1일평균임금": None, "퇴직금추계액": 0.0,
        "기산일": None, "급여기준액": 급여기준액, "배수": 1.0, "임원배수적용": False,
        "warning": None,
    }

    def _add_warning(w):
        result["warning"] = f"{result['warning']} / {w}".strip(" /") if result["warning"] else w

    # 임원 퇴직금 규정상 배수(예: 정관·임원퇴직금지급규정상 2배, 3배 등)는
    # 부서명이 정확히 '임원'인 인원에만 적용한다(일반 직원은 입력값이 있어도 무시).
    if 부서 == EXECUTIVE_DEPT:
        배수_raw = emp.get("배수(임원)")
        배수 = _safe_float(배수_raw, default=1.0)
        if 배수_raw in (None, ""):
            _add_warning("임원 배수 미입력 — 1배로 계산됨")
        elif 배수 <= 0:
            _add_warning("임원 배수 값 오류(0 이하) — 1배로 계산됨")
            배수 = 1.0
        result["배수"] = 배수
        result["임원배수적용"] = True

    if 입사일 is None:
        _add_warning("입사일 미입력 — 재직일수/추계액 계산 불가")
        return result

    기산일 = 입사일
    if 중간정산일 is not None:
        if 중간정산일 < 입사일:
            _add_warning("중간정산일이 입사일보다 이전 — 확인 필요(입사일 기준으로 계산)")
        elif 중간정산일 > 당기결산일:
            _add_warning("중간정산일이 결산기준일 이후 — 확인 필요(입사일 기준으로 계산)")
        else:
            기산일 = 중간정산일

    if 입사일 > 당기결산일:
        _add_warning("입사일이 결산기준일 이후 — 확인 필요('당기정보'에 등재된 인원임)")
        return result

    재직일수 = (당기결산일 - 기산일).days
    result["재직일수"] = 재직일수
    result["기산일"] = 기산일
    일평균임금 = 급여기준액 * 12 / 365
    result["1일평균임금"] = 일평균임금
    result["퇴직금추계액"] = 일평균임금 * 30 * (재직일수 / 365) * result["배수"]
    if 급여기준액 <= 0:
        _add_warning("월평균급여(상여금 포함) 미입력 — 추계액 0으로 계산됨")

    return result


# ── 명세서 구성 ('당기정보' 시트 기준) ────────────────────────────────────

def build_schedule_table(employees: list, 당기결산일: date, 신규입사_keys: set) -> pd.DataFrame:
    rows = []
    for e in employees:
        r = compute_employee(e, 당기결산일)
        원가구분 = _cost_type(e)
        신규입사 = _employee_key(e) in 신규입사_keys
        중간정산일_raw = _safe_date(e.get("중간정산일"))
        중간정산반영 = r["기산일"] is not None and 중간정산일_raw is not None and r["기산일"] == 중간정산일_raw

        비고 = e.get("비고") or ""
        tags = []
        if r["warning"]:
            tags.append(r["warning"])
        if 신규입사:
            tags.append("당기 신규입사(전기정보에는 없음)")
        if 중간정산반영:
            tags.append(f"퇴직금 중간정산 이력 있음(재직일수는 {중간정산일_raw}부터 기산)")
        if r["임원배수적용"]:
            tags.append(f"임원 배수 {r['배수']:g}배 적용")
        if tags:
            비고 = f"{비고} / {' / '.join(tags)}".strip(" /")

        rows.append({
            "사업장": e.get("사업장") or "",
            "부서": e.get("부서") or "",
            "사번": e.get("사번"),
            "성명": e.get("성명"),
            "직급": e.get("직급"),
            "원가구분": 원가구분,
            "입사일": e.get("입사일"),
            "중간정산일": e.get("중간정산일"),
            "월평균급여(원)": _safe_float(e.get("월평균급여(원)")),
            "상여금(연간, 원)": _safe_float(e.get("상여금(연간, 원)")),
            "월평균급여(상여금가산후)(계산)": r["급여기준액"],
            "배수(임원)": r["배수"] if r["임원배수적용"] else None,
            "재직일수(당기말기준)": r["재직일수"],
            "1일평균임금(계산)": r["1일평균임금"],
            "퇴직금추계액(계산)": r["퇴직금추계액"],
            "당기신규입사": 신규입사,
            "중간정산반영": 중간정산반영,
            "비고": 비고,
        })

    df = pd.DataFrame(rows)
    if df.empty:
        return df
    return df.sort_values(["사업장", "부서", "원가구분", "사번"], na_position="last").reset_index(drop=True)


# ── 요약(원가구분별 대사 + 인원변동) ──────────────────────────────────────────

def build_summary(당기_df: pd.DataFrame, 전기_employees: list,
                   신규입사자_recs: list, 퇴사자_recs: list, basis: dict) -> dict:
    if 당기_df.empty and not 전기_employees and not 신규입사자_recs and not 퇴사자_recs:
        return {}

    d = 당기_df.copy()

    by_cost = []
    총_당기재계산 = 0.0
    총_전기입력 = 0.0
    총_당기회사계상 = None
    for cost in COST_TYPES:
        cdf = d[d["원가구분"] == cost] if not d.empty else d
        당기재계산 = float(cdf["퇴직금추계액(계산)"].sum()) if not cdf.empty else 0.0
        전기입력 = basis.get(BASIS_KEYS[cost]["전기"])
        당기회사계상 = basis.get(BASIS_KEYS[cost]["당기"])
        by_cost.append({
            "구분": cost,
            "전기말(회사계상)": 전기입력,
            "당기말(재계산)": 당기재계산,
            "당기말(회사계상)": 당기회사계상,
            "증감(재계산-전기)": 당기재계산 - (전기입력 or 0.0),
            "대사차이(재계산-회사계상)": None if 당기회사계상 is None else 당기재계산 - 당기회사계상,
        })
        총_당기재계산 += 당기재계산
        총_전기입력 += (전기입력 or 0.0)
        if 당기회사계상 is not None:
            총_당기회사계상 = (총_당기회사계상 or 0.0) + 당기회사계상

    미분류 = d[~d["원가구분"].isin(COST_TYPES)] if not d.empty else d
    if not 미분류.empty:
        미분류_재계산 = float(미분류["퇴직금추계액(계산)"].sum())
        by_cost.append({
            "구분": "(미분류)",
            "전기말(회사계상)": None,
            "당기말(재계산)": 미분류_재계산,
            "당기말(회사계상)": None,
            "증감(재계산-전기)": 미분류_재계산,
            "대사차이(재계산-회사계상)": None,
        })
        총_당기재계산 += 미분류_재계산

    total_row = {
        "구분": "합계",
        "전기말(회사계상)": 총_전기입력,
        "당기말(재계산)": 총_당기재계산,
        "당기말(회사계상)": 총_당기회사계상,
        "증감(재계산-전기)": 총_당기재계산 - 총_전기입력,
        "대사차이(재계산-회사계상)": None if 총_당기회사계상 is None else 총_당기재계산 - 총_당기회사계상,
    }

    전기_df = _to_display_df(전기_employees)
    신규입사자_df = _to_display_df(신규입사자_recs)
    퇴사자_df = _to_display_df(퇴사자_recs)

    headcount = {
        "전기말인원수": int(len(전기_df)),
        "신규입사인원수": int(len(신규입사자_df)),
        "퇴사인원수": int(len(퇴사자_df)),
        "당기말인원수": int(len(d)),
    }
    headcount_by_cost = []
    for cost in COST_TYPES:
        headcount_by_cost.append({
            "원가구분": cost,
            "전기말인원수": int(len(전기_df[전기_df["원가구분"] == cost])) if not 전기_df.empty else 0,
            "신규입사인원수": int(len(신규입사자_df[신규입사자_df["원가구분"] == cost])) if not 신규입사자_df.empty else 0,
            "퇴사인원수": int(len(퇴사자_df[퇴사자_df["원가구분"] == cost])) if not 퇴사자_df.empty else 0,
            "당기말인원수": int(len(d[d["원가구분"] == cost])) if not d.empty else 0,
        })

    return {
        "by_cost": by_cost,
        "total_row": total_row,
        "신규입사자": 신규입사자_df[["사업장", "부서", "사번", "성명", "직급", "원가구분", "입사일"]],
        "퇴사자": 퇴사자_df,
        "headcount": headcount,
        "headcount_by_cost": headcount_by_cost,
    }


def write_summary_sheet(ws, summary: dict, company: str, target_fy: str,
                         당기결산일: date, 전기결산일: date, interim_month: int = None):
    header_fill = PatternFill("solid", fgColor="4472C4")
    header_font = Font(bold=True, color="FFFFFF")
    section_fill = PatternFill("solid", fgColor="203864")
    section_font = Font(bold=True, color="FFFFFF", size=12)
    total_fill = PatternFill("solid", fgColor="9DC3E6")
    sig_fill = PatternFill("solid", fgColor="FFFF00")
    bold = Font(bold=True)
    thin = Side(style="thin", color="B7B7B7")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")

    period_note = f", ~{interim_month}월 중간결산(반기 등)" if interim_month else ""
    ws.cell(row=1, column=1,
            value=(f"퇴직급여충당부채 요약표 (회사: {company}, 회계연도: {target_fy}{period_note}, "
                   f"전기말: {전기결산일}, 당기말: {당기결산일})")).font = Font(bold=True, size=13)
    for col in "ABCDEFG":
        ws.column_dimensions[col].width = 18

    r = 3
    if not summary:
        ws.cell(row=r, column=1, value="(인원 데이터 없음)")
        return

    # 1) 원가구분별 전기·당기 비교 + 대사
    ws.cell(row=r, column=1, value="■ 퇴직급여충당부채 원가구분별 전기·당기 비교 (대사)").fill = section_fill
    ws.cell(row=r, column=1).font = section_font
    for c in range(2, 7):
        ws.cell(row=r, column=c).fill = section_fill
    r += 2

    cost_headers = ["구분", "전기말(회사계상)", "당기말(재계산)", "당기말(회사계상)",
                     "증감(재계산-전기)", "대사차이(재계산-회사계상)"]
    for i, h in enumerate(cost_headers, start=1):
        cell = ws.cell(row=r, column=i, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border
    r += 1

    for row_data in summary["by_cost"] + [summary["total_row"]]:
        is_total = row_data["구분"] == "합계"
        cell = ws.cell(row=r, column=1, value=row_data["구분"])
        cell.border = border
        vals = [row_data["전기말(회사계상)"], row_data["당기말(재계산)"],
                row_data["당기말(회사계상)"], row_data["증감(재계산-전기)"],
                row_data["대사차이(재계산-회사계상)"]]
        if is_total:
            cell.font = bold
            cell.fill = total_fill
        for i, v in enumerate(vals, start=2):
            cell = ws.cell(row=r, column=i, value=v)
            cell.border = border
            if v is not None:
                cell.number_format = "#,##0"
            if is_total:
                cell.font = bold
                cell.fill = total_fill
            elif i == 6 and _is_significant(v, row_data["당기말(회사계상)"]):
                cell.fill = sig_fill
        r += 1
    r += 2

    # 2) 인원 변동 요약
    ws.cell(row=r, column=1, value="■ 인원 변동 요약").fill = section_fill
    ws.cell(row=r, column=1).font = section_font
    for c in range(2, 6):
        ws.cell(row=r, column=c).fill = section_fill
    r += 2

    hc_headers = ["구분", "전기말인원수", "신규입사인원수", "퇴사인원수", "당기말인원수"]
    for i, h in enumerate(hc_headers, start=1):
        cell = ws.cell(row=r, column=i, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border
    r += 1

    for hc in summary["headcount_by_cost"]:
        cell = ws.cell(row=r, column=1, value=hc["원가구분"])
        cell.border = border
        for i, k in enumerate(["전기말인원수", "신규입사인원수", "퇴사인원수", "당기말인원수"], start=2):
            cell = ws.cell(row=r, column=i, value=hc[k])
            cell.border = border
        r += 1
    hc = summary["headcount"]
    cell = ws.cell(row=r, column=1, value="합계")
    cell.font = bold
    cell.fill = total_fill
    cell.border = border
    for i, k in enumerate(["전기말인원수", "신규입사인원수", "퇴사인원수", "당기말인원수"], start=2):
        cell = ws.cell(row=r, column=i, value=hc[k])
        cell.border = border
        cell.font = bold
        cell.fill = total_fill
    r += 3

    # 3) 신규입사자 명단
    ws.cell(row=r, column=1,
            value="■ 신규입사자 명단 ('당기정보'에는 있으나 '전기정보'에는 없음)").fill = section_fill
    ws.cell(row=r, column=1).font = section_font
    for c in range(2, 8):
        ws.cell(row=r, column=c).fill = section_fill
    r += 2

    new_headers = ["사업장", "부서", "사번", "성명", "직급", "원가구분", "입사일"]
    for i, h in enumerate(new_headers, start=1):
        cell = ws.cell(row=r, column=i, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border
    r += 1
    신규입사자 = summary["신규입사자"]
    if 신규입사자.empty:
        ws.cell(row=r, column=1, value="(해당 없음)")
        for c in range(1, 8):
            ws.cell(row=r, column=c).border = border
        r += 1
    else:
        for _, row in 신규입사자.iterrows():
            for i, h in enumerate(new_headers, start=1):
                val = row[h]
                cell = ws.cell(row=r, column=i, value=(None if pd.isna(val) else val))
                cell.border = border
                if h == "입사일" and val is not None and not pd.isna(val):
                    cell.number_format = "yyyy-mm-dd"
            r += 1
    r += 2

    # 4) 퇴사자 명단
    ws.cell(row=r, column=1,
            value="■ 퇴사자 명단 ('전기정보'에는 있으나 '당기정보'에는 없음)").fill = section_fill
    ws.cell(row=r, column=1).font = section_font
    for c in range(2, 9):
        ws.cell(row=r, column=c).fill = section_fill
    r += 1
    ws.cell(row=r, column=1,
            value="※ 사번(없으면 성명) 기준 매칭 결과이며, 정확한 퇴사일은 별도 자료가 없는 한 알 수 없어 "
                  "전기 시점 입사일/중간정산일만 참고로 표시합니다.").font = Font(italic=True, color="808080", size=9)
    r += 1

    left_headers = ["사업장", "부서", "사번", "성명", "직급", "원가구분", "입사일", "중간정산일"]
    for i, h in enumerate(left_headers, start=1):
        cell = ws.cell(row=r, column=i, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border
    r += 1
    퇴사자 = summary["퇴사자"]
    if 퇴사자.empty:
        ws.cell(row=r, column=1, value="(해당 없음)")
        for c in range(1, 9):
            ws.cell(row=r, column=c).border = border
        r += 1
    else:
        for _, row in 퇴사자.iterrows():
            for i, h in enumerate(left_headers, start=1):
                val = row[h]
                cell = ws.cell(row=r, column=i, value=(None if pd.isna(val) else val))
                cell.border = border
                if h in ("입사일", "중간정산일") and val is not None and not pd.isna(val):
                    cell.number_format = "yyyy-mm-dd"
            r += 1


# ── 엑셀 저장 ────────────────────────────────────────────────────────────────

MONEY_COLS = ["월평균급여(원)", "상여금(연간, 원)", "월평균급여(상여금가산후)(계산)", "1일평균임금(계산)", "퇴직금추계액(계산)"]
DATE_COLS = ["입사일", "중간정산일"]


def save_results(df: pd.DataFrame, output_path: str, company: str, target_fy: str,
                  당기결산일: date, 전기결산일: date, 전기_employees: list,
                  신규입사자_recs: list, 퇴사자_recs: list, basis: dict, interim_month: int = None):
    wb = openpyxl.Workbook()
    ws_summary = wb.active
    ws_summary.title = "요약표"
    write_summary_sheet(ws_summary, build_summary(df, 전기_employees, 신규입사자_recs, 퇴사자_recs, basis),
                         company, target_fy, 당기결산일, 전기결산일, interim_month)

    ws = wb.create_sheet("인원별추계명세")

    header_fill = PatternFill("solid", fgColor="4472C4")
    header_font = Font(bold=True, color="FFFFFF")
    subtotal_fill = PatternFill("solid", fgColor="D9E1F2")
    total_fill = PatternFill("solid", fgColor="9DC3E6")
    settle_fill = PatternFill("solid", fgColor="FCE4D6")
    bold = Font(bold=True)
    thin = Side(style="thin", color="B7B7B7")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")

    period_note = f", ~{interim_month}월 중간결산(반기 등)" if interim_month else ""
    ws.cell(row=1, column=1,
            value=(f"인원별 퇴직금 추계 명세서 (회사: {company}, 회계연도: {target_fy}{period_note}, "
                   f"당기말 결산기준일: {당기결산일})")).font = Font(bold=True, size=13)

    headers = list(df.columns) if not df.empty else []
    header_row = 3
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=header_row, column=i, value=h)
        c.fill = header_fill
        c.font = header_font
        c.alignment = center
        c.border = border
        ws.column_dimensions[get_column_letter(i)].width = 16 if h != "비고" else 40
    ws.freeze_panes = f"A{header_row + 1}"

    r = header_row + 1
    if df.empty:
        ws.cell(row=r, column=1, value="(인원 데이터 없음)")
        os.makedirs(os.path.dirname(output_path), exist_ok=True)
        wb.save(output_path)
        return

    totals = {c: 0.0 for c in MONEY_COLS}
    grand_totals = {c: 0.0 for c in MONEY_COLS}

    for cost, gdf in df.groupby("원가구분", sort=False):
        for _, row in gdf.iterrows():
            is_settled = bool(row.get("중간정산반영"))
            for i, h in enumerate(headers, start=1):
                val = row[h]
                cell = ws.cell(row=r, column=i, value=(None if pd.isna(val) else val))
                cell.border = border
                if h in DATE_COLS and val is not None and not pd.isna(val):
                    cell.number_format = "yyyy-mm-dd"
                if h in MONEY_COLS:
                    cell.number_format = "#,##0"
                if h in ("당기신규입사", "중간정산반영"):
                    cell.value = "O" if val is True else None
                    cell.alignment = center
                if is_settled and h not in ("비고",):
                    cell.fill = settle_fill
            for c in MONEY_COLS:
                v = row.get(c)
                if v is not None and not pd.isna(v):
                    totals[c] += v
                    grand_totals[c] += v
            r += 1

        ws.cell(row=r, column=1, value=f"[{cost} 소계]").font = bold
        for i, h in enumerate(headers, start=1):
            cell = ws.cell(row=r, column=i)
            cell.fill = subtotal_fill
            cell.border = border
            if h in MONEY_COLS:
                cell.value = totals[h]
                cell.number_format = "#,##0"
                cell.font = bold
        r += 1
        totals = {c: 0.0 for c in MONEY_COLS}

    ws.cell(row=r, column=1, value="총계").font = Font(bold=True, size=11)
    for i, h in enumerate(headers, start=1):
        cell = ws.cell(row=r, column=i)
        cell.fill = total_fill
        cell.border = border
        if h in MONEY_COLS:
            cell.value = grand_totals[h]
            cell.number_format = "#,##0"
            cell.font = bold

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)


# ── 메인 ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="퇴직급여충당부채 검증앱(일반기업회계기준·퇴직금 추계액 방식)")
    parser.add_argument("company", nargs="?", default=None, help="처리할 회사명 (생략 시 파일 자동 탐색)")
    parser.add_argument("--file", default=None, help="처리할 특정 입력 파일명 (input_data/ 기준)")
    parser.add_argument("--fiscal-month", type=int, default=12, help="결산월 (기본 12월). 예: 6월 결산법인이면 6")
    parser.add_argument("--fiscal-year", default=None, help="검증 대상 회계연도 (예: 2026). 생략 시 입력파일명의 fy 뒤 숫자 사용")
    parser.add_argument("--interim-month", type=int, default=None,
                         help="반기 등 중간결산 검토월 (예: 6). 당기말 결산기준일만 해당 월말로 앞당기고, "
                              "전기말은 직전 회계연도 결산기준일 그대로 사용")
    args = parser.parse_args()

    input_path = _find_input_file(args.company, args.file)
    company = args.company or os.path.basename(input_path).split("_")[1]

    target_fy = args.fiscal_year
    if not target_fy:
        base = os.path.basename(input_path)
        idx = base.lower().find("fy")
        if idx == -1:
            raise ValueError("--fiscal-year 를 지정하거나 파일명에 'fy<연도>'를 포함하세요.")
        digits = ""
        for ch in base[idx + 2:]:
            if ch.isdigit():
                digits += ch
            else:
                break
        target_fy = f"20{digits}" if len(digits) == 2 else digits

    _, fy_end_ym = _fy_bounds(target_fy, args.fiscal_month)
    fy_end_ym = _apply_interim(fy_end_ym, target_fy, args.interim_month)
    당기결산일 = _ym_to_end_date(fy_end_ym)

    prior_fy = str(int(target_fy) - 1)
    _, prior_fy_end_ym = _fy_bounds(prior_fy, args.fiscal_month)
    전기결산일 = _ym_to_end_date(prior_fy_end_ym)

    print(f"[입력] {input_path}")
    interim_note = f", 중간결산월={args.interim_month}(반기 등)" if args.interim_month else ""
    print(f"[대상] 회사={company}, 회계연도={target_fy}, 결산월={args.fiscal_month}{interim_note}")
    print(f"[기준일] 당기말={당기결산일}, 전기말={전기결산일}")

    당기_employees = load_employees(input_path, CURRENT_SHEET)
    전기_employees = load_employees(input_path, PRIOR_SHEET)
    basis = load_basis(input_path)
    print(f"[인원 수] 당기={len(당기_employees)}건, 전기={len(전기_employees)}건")

    matched = match_periods(당기_employees, 전기_employees)
    df = build_schedule_table(당기_employees, 당기결산일, matched["신규입사_keys"])

    suffix = f"_interim{args.interim_month:02d}" if args.interim_month else ""
    output_path = os.path.join(OUTPUT_DIR, f"severance_schedule_{company}_{target_fy}{suffix}.xlsx")
    save_results(df, output_path, company, target_fy, 당기결산일, 전기결산일,
                 전기_employees, matched["신규입사자"], matched["퇴사자"], basis, args.interim_month)
    print(f"[완료] {output_path}")


if __name__ == "__main__":
    main()
