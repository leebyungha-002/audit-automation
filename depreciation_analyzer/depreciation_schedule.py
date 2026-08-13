"""유형자산·투자부동산 감가상각비 검증앱 — 엔진.

input_data/depreciation_<company>_information_fy<year>.xlsx 를 읽어
자산별 월별 상각 스케줄을 재계산하고, 계정과목별 소계가 있는
'고정자산명세서'를 output/depreciation_schedule_<company>_<fy>.xlsx 로 생성한다.

회사계상 감가상각누계액/당기상각비를 입력해 두면 앱의 재계산 결과와 자동 대사해
차이를 표시한다(유의차이는 노란색 강조).

실행 예:
    python depreciation_schedule.py kyungnam --fiscal-month 6
    python depreciation_schedule.py --file depreciation_kyungnam_information_fy2026.xlsx --fiscal-month 6
"""
import argparse
import glob
import os

import pandas as pd
import openpyxl
from dateutil.relativedelta import relativedelta
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(HERE, "input_data")
OUTPUT_DIR = os.path.join(HERE, "output")

SIG_THRESHOLD_ABS = 1000      # 유의차이 절대금액 기준(원)
SIG_THRESHOLD_PCT = 0.01      # 유의차이 비율 기준(1%)


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


def _fiscal_year(ym: str, fiscal_month: int) -> str:
    """결산월 기준 회계연도 문자열 반환 (lease_schedule.py와 동일 규칙).
    결산월=12 → 캘린더 연도 그대로.
    결산월=6  → 7월 이후는 다음 연도로 귀속 (예: 2023-07 → '2024')."""
    if fiscal_month == 12:
        return ym[:4]
    year, month = int(ym[:4]), int(ym[5:7])
    return str(year + 1) if month > fiscal_month else str(year)


# ── 입력 로딩 ────────────────────────────────────────────────────────────────

def _find_input_file(company: str = None, file: str = None) -> str:
    if file:
        path = file if os.path.isabs(file) else os.path.join(INPUT_DIR, file)
        if not os.path.exists(path):
            raise FileNotFoundError(f"입력 파일을 찾을 수 없습니다: {path}")
        return path

    if company:
        pattern = os.path.join(INPUT_DIR, f"depreciation_{company}_information_fy*.xlsx")
    else:
        pattern = os.path.join(INPUT_DIR, "depreciation_*_information_fy*.xlsx")

    matches = [p for p in glob.glob(pattern) if "template" not in os.path.basename(p)]
    if not matches:
        raise FileNotFoundError(f"입력 파일을 찾을 수 없습니다: {pattern}")
    if len(matches) > 1 and not company:
        raise ValueError(f"회사를 특정해주세요. 후보 파일 여러 개: {matches}")
    return matches[0]


def load_assets(path: str) -> list:
    """'자산정보' 시트(1~2행 헤더, 3행부터 데이터)를 읽어 dict 목록으로 반환."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["자산정보"] if "자산정보" in wb.sheetnames else wb.worksheets[0]
    headers = [c.value for c in ws[2]]
    assets = []
    for row in ws.iter_rows(min_row=3, values_only=True):
        if row is None or all(v is None for v in row):
            continue
        rec = dict(zip(headers, row))
        if not rec.get("자산명(세부내역)") and not rec.get("자산관리번호"):
            continue
        assets.append(rec)
    return assets


# ── 상각 스케줄 계산 ─────────────────────────────────────────────────────────

def build_asset_schedule(a: dict) -> tuple:
    """단일 자산의 취득일~내용연수(또는 처분일) 전체 월별 상각 스케줄 계산.
    Returns (schedule_df, warning: str|None)
    """
    warning = None
    acquire = _safe_date(a.get("취득일"))
    cost = _safe_float(a.get("취득원가"))
    residual = _safe_float(a.get("잔존가치"))
    life_years = _safe_float(a.get("내용연수(년)"))
    n_months = round(life_years * 12)
    method = str(a.get("상각방법(정액법/정률법)") or "정액법").strip()
    rate = _safe_float(a.get("상각률(정률법전용)"))
    amort_opt = str(a.get("상각개시(당월/익월)") or "당월").strip()
    offset = 1 if amort_opt == "익월" else 0
    dispose = _safe_date(a.get("처분일"))
    impair_date = _safe_date(a.get("손상차손 인식일"))
    impair_amt = _safe_float(a.get("손상차손 인식액"))

    if acquire is None or n_months <= 0 or cost <= 0:
        cols = ["연월", "기초장부가액", "당월상각비", "손상차손인식액", "기말장부가액"]
        return pd.DataFrame(columns=cols), "취득일/취득원가/내용연수 중 필수값 누락 — 상각 계산 불가"

    if method == "정률법" and rate <= 0:
        method = "정액법"
        warning = "상각률 미입력 → 정액법으로 대체 계산"

    # 손상 전(全 기간) 정액법 월상각액. 손상 인식월에 도달하면 잔여 개월수 기준으로 재계산한다
    # (내용연수 자체는 재평가하지 않는다는 전제 — 원래 종료시점은 그대로 유지).
    monthly_straight = (cost - residual) / n_months if n_months > 0 else 0.0
    dispose_ym = dispose.strftime("%Y-%m") if dispose else None
    impair_ym = impair_date.strftime("%Y-%m") if impair_date else None
    impair_applied = False

    rows = []
    book = cost
    for mi in range(1, n_months + 1):
        dt = acquire + relativedelta(months=mi - 1 + offset)
        ym = dt.strftime("%Y-%m")
        if dispose_ym is not None and ym > dispose_ym:
            break

        open_book = book
        if method == "정률법":
            dep = open_book * (rate / 12)
        else:
            dep = monthly_straight
        dep = max(0.0, min(dep, open_book - residual))
        close_book = open_book - dep

        impair_this_month = 0.0
        if impair_ym is not None and not impair_applied and ym == impair_ym and impair_amt > 0:
            # 손상 인식월의 정상 상각을 먼저 반영한 뒤, 그 시점 장부금액에서 손상차손을 차감(0 하한)
            impair_this_month = min(impair_amt, close_book)
            close_book = max(0.0, close_book - impair_this_month)
            impair_applied = True
            remaining_months = n_months - mi
            if method != "정률법":
                monthly_straight = (close_book - residual) / remaining_months if remaining_months > 0 else 0.0

        rows.append({
            "연월": ym, "기초장부가액": open_book, "당월상각비": dep,
            "손상차손인식액": impair_this_month, "기말장부가액": close_book,
        })
        book = close_book
        if book <= residual + 1e-6:
            break

    return pd.DataFrame(rows), warning


def summarize_for_fy(sched: pd.DataFrame, fiscal_month: int, target_fy: str) -> dict:
    """전체 월별 스케줄에서 특정 회계연도(target_fy)의 기초/당기/기말 누계상각액·손상차손누계액 집계."""
    empty = {"기초누계": 0.0, "당기상각비": 0.0, "기말누계": 0.0,
             "기초손상누계": 0.0, "당기손상": 0.0, "기말손상누계": 0.0}
    if sched.empty:
        return empty
    fy_col = sched["연월"].apply(lambda ym: _fiscal_year(ym, fiscal_month))
    기초누계 = sched.loc[fy_col < target_fy, "당월상각비"].sum()
    당기상각비 = sched.loc[fy_col == target_fy, "당월상각비"].sum()
    기초손상 = sched.loc[fy_col < target_fy, "손상차손인식액"].sum() if "손상차손인식액" in sched.columns else 0.0
    당기손상 = sched.loc[fy_col == target_fy, "손상차손인식액"].sum() if "손상차손인식액" in sched.columns else 0.0
    return {
        "기초누계": 기초누계, "당기상각비": 당기상각비, "기말누계": 기초누계 + 당기상각비,
        "기초손상누계": 기초손상, "당기손상": 당기손상, "기말손상누계": 기초손상 + 당기손상,
    }


# ── 명세서 구성 ──────────────────────────────────────────────────────────────

def build_schedule_table(assets: list, fiscal_month: int, target_fy: str) -> pd.DataFrame:
    rows = []
    for a in assets:
        sched, warning = build_asset_schedule(a)
        agg = summarize_for_fy(sched, fiscal_month, target_fy)
        cost = _safe_float(a.get("취득원가"))
        기말장부가액 = cost - agg["기말누계"] - agg["기말손상누계"]

        company_beg = a.get("전기말 회사계상 감가상각누계액")
        company_dep = a.get("당기 회사계상 감가상각비")
        company_beg_f = _safe_float(company_beg) if company_beg not in (None, "") else None
        company_dep_f = _safe_float(company_dep) if company_dep not in (None, "") else None

        비고 = a.get("비고") or ""
        if warning:
            비고 = f"{비고} / {warning}".strip(" /")

        rows.append({
            "사업장": a.get("사업장") or "",
            "자산분류": a.get("자산분류(유형자산/투자부동산/무형자산)") or "",
            "계정과목": a.get("계정과목") or "(미분류)",
            "자산관리번호": a.get("자산관리번호"),
            "자산명(세부내역)": a.get("자산명(세부내역)"),
            "취득일": a.get("취득일"),
            "취득원가": cost,
            "기초감가상각누계액(계산)": agg["기초누계"],
            "당기감가상각비(계산)": agg["당기상각비"],
            "기말감가상각누계액(계산)": agg["기말누계"],
            "전기말 손상차손누계액(계산)": agg["기초손상누계"],
            "당기 손상차손인식액(계산)": agg["당기손상"],
            "기말 손상차손누계액(계산)": agg["기말손상누계"],
            "기말장부가액(계산)": 기말장부가액,
            "전기말 회사계상 누계액": company_beg_f,
            "기초누계액 차이": (None if company_beg_f is None else agg["기초누계"] - company_beg_f),
            "당기 회사계상 상각비": company_dep_f,
            "당기상각비 차이": (None if company_dep_f is None else agg["당기상각비"] - company_dep_f),
            "원가구분": a.get("원가구분"),
            "비고": 비고,
        })

    df = pd.DataFrame(rows)
    if df.empty:
        return df
    return df.sort_values(["사업장", "계정과목", "자산명(세부내역)"], na_position="last").reset_index(drop=True)


# ── 계정분류별 요약표 ─────────────────────────────────────────────────────────

CATEGORY_ORDER = ["유형자산", "투자부동산", "무형자산"]


def build_category_summary(df: pd.DataFrame) -> dict:
    """자산분류(유형자산/투자부동산/무형자산)별 요약 지표 + 사업장×원가구분 감가상각비 피벗 계산."""
    if df.empty:
        return {}

    d = df.copy()
    d["자산분류"] = d["자산분류"].astype(str).str.strip().replace({"": "(미분류)", "nan": "(미분류)"})
    d["사업장"] = d["사업장"].astype(str).str.strip().replace({"": "(미분류)", "nan": "(미분류)"})
    d["원가구분"] = d["원가구분"].astype(str).str.strip().replace({"": "(미분류)", "None": "(미분류)", "nan": "(미분류)"})
    d["기초장부금액"] = d["취득원가"] - d["기초감가상각누계액(계산)"] - d["전기말 손상차손누계액(계산)"]

    present = list(dict.fromkeys(d["자산분류"]))
    ordered = [c for c in CATEGORY_ORDER if c in present] + [c for c in present if c not in CATEGORY_ORDER]

    summaries = {}
    for cat in ordered:
        cdf = d[d["자산분류"] == cat]
        pivot = cdf.pivot_table(
            index="사업장", columns="원가구분", values="당기감가상각비(계산)",
            aggfunc="sum", fill_value=0.0,
        )
        pivot["소계"] = pivot.sum(axis=1)
        grand = pivot.sum(axis=0)
        grand.name = "총계"
        pivot = pd.concat([pivot, grand.to_frame().T])

        summaries[cat] = {
            "자산수": int(len(cdf)),
            "기초장부금액": cdf["기초장부금액"].sum(),
            "당기감가상각비": cdf["당기감가상각비(계산)"].sum(),
            "당기손상차손": cdf["당기 손상차손인식액(계산)"].sum(),
            "기말장부금액": cdf["기말장부가액(계산)"].sum(),
            "pivot": pivot,
        }
    return summaries


def write_summary_sheet(ws, summaries: dict, company: str, target_fy: str):
    header_fill = PatternFill("solid", fgColor="4472C4")
    header_font = Font(bold=True, color="FFFFFF")
    section_fill = PatternFill("solid", fgColor="203864")
    section_font = Font(bold=True, color="FFFFFF", size=12)
    total_fill = PatternFill("solid", fgColor="9DC3E6")
    bold = Font(bold=True)
    thin = Side(style="thin", color="B7B7B7")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")

    ws.cell(row=1, column=1, value=f"고정자산 계정분류별 요약표 (회사: {company}, 회계연도: {target_fy})").font = Font(bold=True, size=13)
    ws.column_dimensions["A"].width = 16
    for col in "BCDEFGH":
        ws.column_dimensions[col].width = 16

    r = 3
    if not summaries:
        ws.cell(row=r, column=1, value="(자산 데이터 없음)")
        return

    METRIC_COLS = ["상각대상 자산수", "기초장부금액", "당기감가상각비", "당기손상차손", "기말장부금액"]

    for cat, s in summaries.items():
        ws.cell(row=r, column=1, value=f"■ {cat}").fill = section_fill
        ws.cell(row=r, column=1).font = section_font
        for c in range(2, 8):
            ws.cell(row=r, column=c).fill = section_fill
        r += 2

        for i, h in enumerate(METRIC_COLS, start=1):
            cell = ws.cell(row=r, column=i, value=h)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center
            cell.border = border
        r += 1
        values = [s["자산수"], s["기초장부금액"], s["당기감가상각비"], s["당기손상차손"], s["기말장부금액"]]
        for i, v in enumerate(values, start=1):
            cell = ws.cell(row=r, column=i, value=v)
            cell.border = border
            if i > 1:
                cell.number_format = "#,##0"
        r += 2

        ws.cell(row=r, column=1, value="사업장별 당기감가상각비 (원가구분별)").font = bold
        r += 1
        pivot = s["pivot"]
        cost_cols = list(pivot.columns)
        headers = ["사업장"] + cost_cols
        for i, h in enumerate(headers, start=1):
            cell = ws.cell(row=r, column=i, value=h)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center
            cell.border = border
        r += 1
        for site, row in pivot.iterrows():
            is_total = site == "총계"
            cell = ws.cell(row=r, column=1, value=site)
            cell.border = border
            if is_total:
                cell.font = bold
                cell.fill = total_fill
            for i, c in enumerate(cost_cols, start=2):
                cell = ws.cell(row=r, column=i, value=row[c])
                cell.number_format = "#,##0"
                cell.border = border
                if is_total:
                    cell.font = bold
                    cell.fill = total_fill
            r += 1
        r += 2


# ── 엑셀 저장 ────────────────────────────────────────────────────────────────

MONEY_COLS = [
    "취득원가", "기초감가상각누계액(계산)", "당기감가상각비(계산)", "기말감가상각누계액(계산)",
    "전기말 손상차손누계액(계산)", "당기 손상차손인식액(계산)", "기말 손상차손누계액(계산)",
    "기말장부가액(계산)", "전기말 회사계상 누계액", "기초누계액 차이",
    "당기 회사계상 상각비", "당기상각비 차이",
]


def _is_significant(diff, base) -> bool:
    if diff is None or pd.isna(diff):
        return False
    if abs(diff) >= SIG_THRESHOLD_ABS and (base in (None, 0) or abs(diff) >= abs(base) * SIG_THRESHOLD_PCT):
        return True
    return False


def save_results(df: pd.DataFrame, output_path: str, company: str, target_fy: str):
    wb = openpyxl.Workbook()
    ws_summary = wb.active
    ws_summary.title = "요약표"
    write_summary_sheet(ws_summary, build_category_summary(df), company, target_fy)

    ws = wb.create_sheet("고정자산명세서")

    header_fill = PatternFill("solid", fgColor="4472C4")
    header_font = Font(bold=True, color="FFFFFF")
    subtotal_fill = PatternFill("solid", fgColor="D9E1F2")
    total_fill = PatternFill("solid", fgColor="9DC3E6")
    sig_fill = PatternFill("solid", fgColor="FFFF00")
    bold = Font(bold=True)
    thin = Side(style="thin", color="B7B7B7")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center")

    ws.cell(row=1, column=1, value=f"고정자산명세서 (회사: {company}, 회계연도: {target_fy})").font = Font(bold=True, size=13)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(df.columns) if not df.empty else 15)

    headers = list(df.columns)
    header_row = 3
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=header_row, column=i, value=h)
        c.fill = header_fill
        c.font = header_font
        c.alignment = center
        c.border = border
        ws.column_dimensions[get_column_letter(i)].width = 16 if h not in ("자산명(세부내역)", "비고") else 26
    ws.freeze_panes = f"A{header_row + 1}"

    r = header_row + 1
    totals = {c: 0.0 for c in MONEY_COLS}
    site_totals = {c: 0.0 for c in MONEY_COLS}
    grand_totals = {c: 0.0 for c in MONEY_COLS}

    if df.empty:
        ws.cell(row=r, column=1, value="(자산 데이터 없음)")
        wb.save(output_path)
        return

    # 사업장 컬럼에 실제 값이 하나라도 있으면 사업장→계정과목 2단 소계, 없으면 기존처럼 계정과목 소계만
    has_site = "사업장" in df.columns and (df["사업장"].astype(str).str.strip() != "").any()
    if has_site:
        df = df.copy()
        df["사업장"] = df["사업장"].astype(str).str.strip().replace("", "(미분류)")
        outer_groups = list(df.groupby("사업장", sort=False))
    else:
        outer_groups = [(None, df)]

    for site, site_df in outer_groups:
        for account, gdf in site_df.groupby("계정과목", sort=False):
            for _, row in gdf.iterrows():
                for i, h in enumerate(headers, start=1):
                    val = row[h]
                    cell = ws.cell(row=r, column=i, value=(None if pd.isna(val) else val))
                    cell.border = border
                    if h == "취득일" and val is not None and not pd.isna(val):
                        cell.number_format = "yyyy-mm-dd"
                    if h in MONEY_COLS:
                        cell.number_format = "#,##0"
                    if h in ("기초누계액 차이", "당기상각비 차이"):
                        base_col = "전기말 회사계상 누계액" if h == "기초누계액 차이" else "당기 회사계상 상각비"
                        if _is_significant(val, row.get(base_col)):
                            cell.fill = sig_fill
                for c in MONEY_COLS:
                    v = row.get(c)
                    if v is not None and not pd.isna(v):
                        totals[c] += v
                        site_totals[c] += v
                        grand_totals[c] += v
                r += 1

            # 계정과목 소계 행
            ws.cell(row=r, column=1, value=f"[{account} 소계]").font = bold
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

        if has_site:
            # 사업장 소계 행
            ws.cell(row=r, column=1, value=f"[{site} 사업장 소계]").font = bold
            for i, h in enumerate(headers, start=1):
                cell = ws.cell(row=r, column=i)
                cell.fill = total_fill
                cell.border = border
                if h in MONEY_COLS:
                    cell.value = site_totals[h]
                    cell.number_format = "#,##0"
                    cell.font = bold
            r += 1
            site_totals = {c: 0.0 for c in MONEY_COLS}

    # 총계 행
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
    parser = argparse.ArgumentParser(description="유형자산/투자부동산 감가상각비 검증앱")
    parser.add_argument("company", nargs="?", default=None, help="처리할 회사명 (생략 시 파일 자동 탐색)")
    parser.add_argument("--file", default=None, help="처리할 특정 입력 파일명 (input_data/ 기준)")
    parser.add_argument("--fiscal-month", type=int, default=12, help="결산월 (기본 12월). 예: 6월 결산이면 6")
    parser.add_argument("--fiscal-year", default=None, help="검증 대상 회계연도 (예: 2026). 생략 시 입력파일명의 fy 뒤 숫자 사용")
    args = parser.parse_args()

    input_path = _find_input_file(args.company, args.file)
    company = args.company or os.path.basename(input_path).split("_")[1]

    target_fy = args.fiscal_year
    if not target_fy:
        base = os.path.basename(input_path)
        marker = "fy"
        idx = base.lower().find(marker)
        if idx == -1:
            raise ValueError("--fiscal-year 를 지정하거나 파일명에 'fy<연도>'를 포함하세요.")
        digits = ""
        for ch in base[idx + 2:]:
            if ch.isdigit():
                digits += ch
            else:
                break
        target_fy = f"20{digits}" if len(digits) == 2 else digits

    print(f"[입력] {input_path}")
    print(f"[대상] 회사={company}, 회계연도={target_fy}, 결산월={args.fiscal_month}")

    assets = load_assets(input_path)
    print(f"[자산 수] {len(assets)}건")

    df = build_schedule_table(assets, args.fiscal_month, target_fy)

    output_path = os.path.join(OUTPUT_DIR, f"depreciation_schedule_{company}_{target_fy}.xlsx")
    save_results(df, output_path, company, target_fy)
    print(f"[완료] {output_path}")


if __name__ == "__main__":
    main()
