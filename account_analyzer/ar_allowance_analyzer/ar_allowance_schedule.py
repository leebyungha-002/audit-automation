"""매출채권 연령분석 및 대손충당금 설정 검증앱 — 엔진.

input_data/ar_allowance_<company>_information_fy<year>.xlsx 를 읽어 거래처(채권)별 경과일수·
연령구간을 재계산하고, 연령구간별 대손율을 적용한 대손충당금을 재산출해 회사계상액과 대사하는
output/ar_allowance_schedule_<company>_<fy>.xlsx 를 생성한다.

핵심 설계 원칙(감사절차 관점)
1. 개별평가와 집합평가(연령분석)를 반드시 분리한다. 부도·회생절차·소송 등 손상 징후가 있는 거래처는
   연령대와 무관하게 연령분석 모집단에서 빠져야 하며, 채권잔액-개별평가 회수가능예상액으로 별도 계산한다.
   이 분리를 빠뜨리는 것이 실무에서 가장 흔한 오류다.
2. 특수관계자채권은 신용위험 성격이 일반 매출채권과 달라 집합평가 모집단에서 제외하고 별도 표시한다.
3. 상장사(roll rate법)와 비상장사(연령별 대손율 직접설정) 모두 "연령구간별 순채권액×대손율=대손충당금"
   이라는 계산 구조 자체는 같다 — 다른 것은 그 대손율을 구하는 방법(과거 이동매트릭스 통계 추정 vs 실무
   관행상 직접 설정)뿐이다. 이 앱은 그 대손율 산출 과정(특히 roll rate의 이동매트릭스 추정) 자체를
   재현하지 않는다 — K-IFRS 계리보고서 앱과 동일한 이유(통계 모델을 근사 복제하면 차이가 오류인지 모델
   단순화 때문인지 구분할 수 없어 위험)로, 회사(또는 계리/컨설팅)가 이미 산출한 연령별 대손율을 입력받아
   (a) 적용 산식이 맞는지, (b) 연령구간이 커질수록 대손율이 낮아지는 등 비정상 패턴이 없는지, (c) 비상장사는
   최근 실제대손율과 괴리가 큰지를 검증하는 데 스코프를 한정한다.
4. 전기 대비 대손충당금 설정률 변동을 자동 비교하고, 대손충당금 T계정(기초+전입-환입-직접상각=기말)
   tie-out을 다른 계정 검증앱과 동일한 방식으로 제공한다.

실행 예:
    python ar_allowance_schedule.py kyungnam --fiscal-month 12
    python ar_allowance_schedule.py --file ar_allowance_kyungnam_information_fy2026.xlsx
"""
import argparse
import calendar
import glob
import os
from datetime import date, timedelta

import pandas as pd
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.datetime import from_excel

HERE = os.path.dirname(os.path.abspath(__file__))
INPUT_DIR = os.path.join(HERE, "input_data")
OUTPUT_DIR = os.path.join(HERE, "output")

SIG_THRESHOLD_ABS = 1000       # 유의차이 절대금액 기준(원)
SIG_THRESHOLD_PCT = 0.01       # 유의차이 비율 기준(1%)
SIG_THRESHOLD_DAYS_ABS = 1     # 유의차이 절대일수 기준(일)
SIG_THRESHOLD_RATE_PP = 0.01   # 설정률/대손율 유의차이 기준(1%p)

RECEIVABLE_SHEET = "매출채권명세"
RATE_SHEET = "연령별대손율"

DEFAULT_BUCKET_THRESHOLDS = [30, 60, 90, 180, 365]

BASIS_LISTED_LABEL = "상장구분(상장/비상장)"
BASIS_THRESHOLDS_LABEL = "연령구간 상한(일, 콤마구분)"
BASIS_FORWARD_LOOKING_LABEL = "Forward-looking(미래전망정보) 조정 반영 여부(Y/N, 상장사 참고)"
BASIS_PRIOR_AR_LABEL = "전기말 회사계상 매출채권 총액(원)"
BASIS_PRIOR_ALLOWANCE_LABEL = "전기말 회사계상 대손충당금(원)"
BASIS_CURRENT_ALLOWANCE_LABEL = "당기말 회사계상 대손충당금(원)"
BASIS_TRANSFER_IN_LABEL = "당기 대손충당금 전입액(손익, 선택, 원)"
BASIS_REVERSAL_LABEL = "당기 대손충당금 환입액(선택, 원)"
BASIS_WRITEOFF_LABEL = "당기 대손금 직접상각(제각)액(선택, 분개장 기준, 원)"

CATEGORY_POOLED = "집합평가(연령분석)"
CATEGORY_INDIVIDUAL = "개별평가"
CATEGORY_RELATED = "특수관계자(별도검토)"

FORMULA_NOTE_LINES = [
    ("핵심 계산식", True),
    ("경과일수 = 결산기준일 − 연령산정 기산일(결제기일/만기일)", False),
    ("순채권액 = 채권잔액 − 담보/보증 등 차감액", False),
    (f"[{CATEGORY_POOLED}] 대손충당금(계산) = 순채권액 × 해당 연령구간의 회사설정 대손율", False),
    (f"[{CATEGORY_INDIVIDUAL}] 대손충당금(계산) = 순채권액 − 개별평가 회수가능예상액"
     "(미입력 시 순채권액 전액을 잠정 계상하고 경고)", False),
    (f"[{CATEGORY_RELATED}] 신용위험 성격이 달라 위 계산에서 제외, 별도 표로만 표시(대손충당금 별도 검토 필요)", False),
    ("", False),
    ("※ 상장사(roll rate법)·비상장사(연령별 대손율 직접설정) 모두 위 계산 구조는 동일 — 대손율을 구하는", True),
    ("  방법(과거 이동매트릭스 추정 vs 실무관행상 직접설정)만 다르며, 이 앱은 그 산출 과정 자체는 재현하지", False),
    ("  않고 입력된 대손율의 적용과 합리성만 검증한다(roll rate 이동매트릭스 재현은 스코프 제외).", False),
]


# ── 공용 헬퍼 (다른 account_analyzer 모듈과 동일 규칙) ────────────────────────

def _safe_float(v, default: float = 0.0):
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
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        # 셀 서식이 날짜로 인식 안 돼 엑셀 일련번호 그대로 반환되는 경우, pd.Timestamp(숫자)로 바로
        # 변환하면 나노초 단위로 오인식해 1970년 근처로 잘못 변환되므로 반드시 엑셀 기준일로 변환한다.
        try:
            return from_excel(v).date()
        except Exception:
            return None
    ts = pd.Timestamp(v)
    return ts.date() if not pd.isna(ts) else None


def _fy_bounds(target_fy: str, fiscal_month: int) -> tuple:
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
        pattern = os.path.join(INPUT_DIR, f"ar_allowance_{company}_information_fy*.xlsx")
    else:
        pattern = os.path.join(INPUT_DIR, "ar_allowance_*_information_fy*.xlsx")

    matches = [p for p in glob.glob(pattern) if "template" not in os.path.basename(p)]
    if not matches:
        raise FileNotFoundError(f"입력 파일을 찾을 수 없습니다: {pattern}")
    if len(matches) > 1 and not company:
        raise ValueError(f"회사를 특정해주세요. 후보 파일 여러 개: {matches}")
    return matches[0]


def _is_significant(diff, base, abs_threshold: float = SIG_THRESHOLD_ABS) -> bool:
    if diff is None or pd.isna(diff):
        return False
    if abs(diff) >= abs_threshold and (base in (None, 0) or abs(diff) >= abs(base) * SIG_THRESHOLD_PCT):
        return True
    return False


# ── 연령구간 ─────────────────────────────────────────────────────────────────

def parse_bucket_thresholds(basis: dict) -> list:
    raw = basis.get(BASIS_THRESHOLDS_LABEL)
    if raw is None or not str(raw).strip():
        return list(DEFAULT_BUCKET_THRESHOLDS)
    try:
        vals = sorted({int(float(str(x).strip())) for x in str(raw).split(",") if str(x).strip()})
        return vals if vals else list(DEFAULT_BUCKET_THRESHOLDS)
    except ValueError:
        return list(DEFAULT_BUCKET_THRESHOLDS)


def bucket_labels(thresholds: list) -> list:
    labels = ["정상(미도래)"]
    prev = 0
    for t in thresholds:
        labels.append(f"{prev + 1}~{t}일")
        prev = t
    labels.append(f"{thresholds[-1]}일 초과")
    return labels


def bucket_for_days(days, thresholds: list, labels: list) -> str:
    if days is None:
        return "(기산일 미입력)"
    if days <= 0:
        return labels[0]
    for i, t in enumerate(thresholds):
        if days <= t:
            return labels[i + 1]
    return labels[-1]


# ── 입력 로딩 ────────────────────────────────────────────────────────────────

def load_receivables(path: str) -> list:
    wb = openpyxl.load_workbook(path, data_only=True)
    if RECEIVABLE_SHEET not in wb.sheetnames:
        return []
    ws = wb[RECEIVABLE_SHEET]
    headers = [c.value for c in ws[2]]
    records = []
    for row in ws.iter_rows(min_row=3, values_only=True):
        if row is None or all(v is None for v in row):
            continue
        rec = dict(zip(headers, row))
        if not rec.get("거래처명"):
            continue
        records.append(rec)
    return records


def load_rate_table(path: str) -> dict:
    wb = openpyxl.load_workbook(path, data_only=True)
    if RATE_SHEET not in wb.sheetnames:
        return {}
    ws = wb[RATE_SHEET]
    table = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or row[0] is None or str(row[0]).strip() == "":
            continue
        label = str(row[0]).strip()
        rate_raw = row[1] if len(row) > 1 else None
        actual_raw = row[2] if len(row) > 2 else None
        rate = _safe_float(rate_raw) / 100.0 if rate_raw not in (None, "") else None
        actual = _safe_float(actual_raw) / 100.0 if actual_raw not in (None, "") else None
        table[label] = {"rate": rate, "actual_rate": actual}
    return table


def load_basis(path: str) -> dict:
    """'기준정보' 시트를 라벨(A열) 기준으로 읽어 {라벨: 원본값} dict 반환."""
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
        basis[label] = value
    return basis


def _basis_float(basis: dict, label: str):
    v = basis.get(label)
    return None if v in (None, "") else _safe_float(v)


def listed_type(basis: dict) -> str:
    raw = str(basis.get(BASIS_LISTED_LABEL) or "").strip()
    return raw if raw in ("상장", "비상장") else "비상장"


def forward_looking_applied(basis: dict) -> str:
    return str(basis.get(BASIS_FORWARD_LOOKING_LABEL) or "").strip().upper() or "N"


# ── 거래처(채권)별 계산 ───────────────────────────────────────────────────────

def compute_receivable(rec: dict, 결산일: date, thresholds: list, labels: list, rate_table: dict) -> dict:
    채권잔액 = _safe_float(rec.get("채권잔액(원)"))
    담보차감 = _safe_float(rec.get("담보/보증 등 차감액(원)"))
    순채권액 = 채권잔액 - 담보차감
    기산일 = _safe_date(rec.get("연령산정 기산일(결제기일/만기일)"))
    특수관계자 = str(rec.get("특수관계자여부(Y/N)") or "").strip().upper() == "Y"
    개별평가대상 = str(rec.get("개별평가대상여부(Y/N)") or "").strip().upper() == "Y"

    회사계산경과일수_raw = rec.get("회사계산 경과일수(선택,참고용)")
    회사계산경과일수 = _safe_float(회사계산경과일수_raw) if 회사계산경과일수_raw not in (None, "") else None

    warning = None

    def _warn(w):
        nonlocal warning
        warning = f"{warning} / {w}".strip(" /") if warning else w

    if 기산일 is None:
        _warn("연령산정 기산일 미입력 — 경과일수 계산 불가")
        경과일수 = None
    else:
        경과일수 = (결산일 - 기산일).days

    연령구간 = bucket_for_days(경과일수, thresholds, labels)
    경과일수차이 = (경과일수 - 회사계산경과일수) if (경과일수 is not None and 회사계산경과일수 is not None) else None

    if 특수관계자:
        평가구분 = CATEGORY_RELATED
        적용대손율 = None
        대손충당금 = None
    elif 개별평가대상:
        평가구분 = CATEGORY_INDIVIDUAL
        적용대손율 = None
        회수가능액_raw = rec.get("개별평가 회수가능예상액(선택,원)")
        회수가능액 = _safe_float(회수가능액_raw) if 회수가능액_raw not in (None, "") else None
        if 회수가능액 is None:
            _warn("개별평가 회수가능예상액 미입력 — 순채권액 전액을 잠정 대손충당금으로 계상")
            대손충당금 = max(순채권액, 0.0)
        else:
            대손충당금 = max(순채권액 - 회수가능액, 0.0)
    else:
        평가구분 = CATEGORY_POOLED
        rate_info = rate_table.get(연령구간)
        if not rate_info or rate_info.get("rate") is None:
            _warn(f"'{연령구간}' 구간의 대손율이 '{RATE_SHEET}' 시트에 없음 — 대손충당금 0으로 계산")
            적용대손율 = None
            대손충당금 = 0.0
        else:
            적용대손율 = rate_info["rate"]
            대손충당금 = 순채권액 * 적용대손율

    회사계상_raw = rec.get("거래처별 회사계상 대손충당금(원)")
    회사계상 = _safe_float(회사계상_raw) if 회사계상_raw not in (None, "") else None
    차이 = None if (대손충당금 is None or 회사계상 is None) else 대손충당금 - 회사계상

    return {
        "거래처명": rec.get("거래처명"),
        "거래처코드": rec.get("거래처코드(선택)"),
        "특수관계자여부": "Y" if 특수관계자 else "N",
        "채권잔액(원)": 채권잔액,
        "담보/보증 등 차감액(원)": 담보차감,
        "순채권액(원)": 순채권액,
        "연령산정 기산일": 기산일,
        "경과일수(계산)": 경과일수,
        "회사계산 경과일수(선택)": 회사계산경과일수,
        "경과일수차이(계산-회사계산)": 경과일수차이,
        "연령구간": 연령구간,
        "평가구분": 평가구분,
        "적용대손율(%)": None if 적용대손율 is None else 적용대손율 * 100.0,
        "대손충당금(계산)": 대손충당금,
        "거래처별 회사계상 대손충당금(원)": 회사계상,
        "차이(계산-회사계상)": 차이,
        "비고": (f"{rec.get('비고') or ''} / {warning}".strip(" /") if warning else (rec.get("비고") or "")),
    }


def build_detail_table(receivables: list, 결산일: date, thresholds: list, labels: list, rate_table: dict) -> pd.DataFrame:
    rows = [compute_receivable(r, 결산일, thresholds, labels, rate_table) for r in receivables]
    df = pd.DataFrame(rows)
    if df.empty:
        return df
    return df.sort_values(["평가구분", "거래처명"], na_position="last").reset_index(drop=True)


def build_bucket_summary(df: pd.DataFrame, labels: list, rate_table: dict) -> list:
    """연령구간별(집합평가 대상만) 채권잔액·대손율·대손충당금 요약. 대손율이 구간이 커질수록
    감소하는(비정상) 패턴과, 회사설정율이 최근 실제대손율보다 낮은(과소설정 의심) 경우를 경고한다."""
    pooled = df[df["평가구분"] == CATEGORY_POOLED] if not df.empty else df
    rows = []
    prev_rate = None
    for label in labels:
        sub = pooled[pooled["연령구간"] == label] if not pooled.empty else pooled
        순채권액합계 = float(sub["순채권액(원)"].sum()) if not sub.empty else 0.0
        대손충당금합계 = float(sub["대손충당금(계산)"].sum()) if not sub.empty else 0.0
        rate_info = rate_table.get(label, {})
        적용대손율 = rate_info.get("rate")
        실제대손율 = rate_info.get("actual_rate")

        tags = []
        if 적용대손율 is None:
            tags.append(f"⚠ '{RATE_SHEET}' 시트에 이 구간 대손율 없음")
        else:
            if prev_rate is not None and 적용대손율 < prev_rate - 1e-9:
                tags.append("⚠ 이전(더 짧은) 연령구간보다 대손율이 낮음 — 비정상 패턴 의심")
            prev_rate = 적용대손율
            if 실제대손율 is not None and 적용대손율 < 실제대손율 - SIG_THRESHOLD_RATE_PP:
                tags.append("⚠ 회사설정 대손율이 최근 실제대손율보다 낮음 — 과소설정 가능성")

        rows.append({
            "연령구간": label,
            "건수": int(len(sub)),
            "순채권액(원)": 순채권액합계,
            "회사설정 대손율(%)": None if 적용대손율 is None else 적용대손율 * 100.0,
            "최근 실제대손율(참고,%)": None if 실제대손율 is None else 실제대손율 * 100.0,
            "대손율차이(회사설정-실제,%p)": (None if 적용대손율 is None or 실제대손율 is None
                                       else (적용대손율 - 실제대손율) * 100.0),
            "대손충당금(계산,원)": 대손충당금합계,
            "비고": " / ".join(tags),
        })
    return rows


def build_category_summary(df: pd.DataFrame) -> list:
    rows = []
    for cat in (CATEGORY_POOLED, CATEGORY_INDIVIDUAL, CATEGORY_RELATED):
        sub = df[df["평가구분"] == cat] if not df.empty else df
        순채권액합계 = float(sub["순채권액(원)"].sum()) if not sub.empty else 0.0
        대손충당금합계 = float(sub["대손충당금(계산)"].sum()) if not sub.empty and cat != CATEGORY_RELATED else None
        rows.append({
            "평가구분": cat,
            "건수": int(len(sub)),
            "순채권액(원)": 순채권액합계,
            "대손충당금(계산,원)": 대손충당금합계,
            "비고": "집합평가·개별평가 합계에서 제외 — 신용위험 성격이 달라 별도 검토 필요" if cat == CATEGORY_RELATED else "",
        })
    return rows


def build_overall_summary(df: pd.DataFrame, basis: dict) -> dict:
    pooled_individual = df[df["평가구분"].isin([CATEGORY_POOLED, CATEGORY_INDIVIDUAL])] if not df.empty else df
    총순채권액 = float(pooled_individual["순채권액(원)"].sum()) if not pooled_individual.empty else 0.0
    총대손충당금_계산 = float(pooled_individual["대손충당금(계산)"].sum()) if not pooled_individual.empty else 0.0

    당기말_회사계상 = _basis_float(basis, BASIS_CURRENT_ALLOWANCE_LABEL)
    당기말차이 = None if 당기말_회사계상 is None else 총대손충당금_계산 - 당기말_회사계상

    설정률_당기 = (총대손충당금_계산 / 총순채권액) if 총순채권액 else None

    전기말_매출채권 = _basis_float(basis, BASIS_PRIOR_AR_LABEL)
    전기말_충당금 = _basis_float(basis, BASIS_PRIOR_ALLOWANCE_LABEL)
    설정률_전기 = (전기말_충당금 / 전기말_매출채권) if (전기말_매출채권 and 전기말_충당금 is not None) else None

    설정률차이 = None if (설정률_당기 is None or 설정률_전기 is None) else 설정률_당기 - 설정률_전기
    설정률_유의변동 = (설정률차이 is not None and abs(설정률차이) >= SIG_THRESHOLD_RATE_PP)

    당기전입액 = _basis_float(basis, BASIS_TRANSFER_IN_LABEL)
    당기환입액 = _basis_float(basis, BASIS_REVERSAL_LABEL) or 0.0
    당기직접상각액 = _basis_float(basis, BASIS_WRITEOFF_LABEL) or 0.0

    tie_out = None
    if 전기말_충당금 is not None and 당기전입액 is not None and 당기말_회사계상 is not None:
        계산상기말 = 전기말_충당금 + 당기전입액 - 당기환입액 - 당기직접상각액
        tie_out = {
            "전기말(회사계상)": 전기말_충당금,
            "당기 전입액(입력)": 당기전입액,
            "당기 환입액(입력)": 당기환입액,
            "당기 직접상각액(입력)": 당기직접상각액,
            "계산상 당기말": 계산상기말,
            "당기말(회사계상)": 당기말_회사계상,
            "차이(계산상당기말-회사계상)": 계산상기말 - 당기말_회사계상,
        }

    return {
        "총순채권액(집합+개별,원)": 총순채권액,
        "총대손충당금(계산,원)": 총대손충당금_계산,
        "당기말(회사계상,원)": 당기말_회사계상,
        "당기말차이(계산-회사계상,원)": 당기말차이,
        "설정률(당기,재계산)": 설정률_당기,
        "설정률(전기,입력값기준)": 설정률_전기,
        "설정률차이(당기-전기,%p)": None if 설정률차이 is None else 설정률차이 * 100.0,
        "설정률_유의변동": 설정률_유의변동,
        "tie_out": tie_out,
    }


def check_rate_label_mismatch(labels: list, rate_table: dict) -> list:
    warnings = []
    missing = [l for l in labels if l not in rate_table or rate_table[l].get("rate") is None]
    extra = [l for l in rate_table if l not in labels]
    if missing:
        warnings.append(f"⚠ '{RATE_SHEET}' 시트에 대손율이 없는 연령구간: {', '.join(missing)}")
    if extra:
        warnings.append(f"⚠ '{RATE_SHEET}' 시트에 있으나 현재 연령구간 설정과 맞지 않는(사용되지 않는) 구간: {', '.join(extra)}")
    return warnings


# ── 결과 저장 ────────────────────────────────────────────────────────────────

DETAIL_MONEY_COLS = ["채권잔액(원)", "담보/보증 등 차감액(원)", "순채권액(원)", "대손충당금(계산)",
                      "거래처별 회사계상 대손충당금(원)", "차이(계산-회사계상)"]
DETAIL_DATE_COLS = ["연령산정 기산일"]
DETAIL_DAY_COLS = ["경과일수(계산)", "회사계산 경과일수(선택)", "경과일수차이(계산-회사계산)"]
DETAIL_PCT_COLS = ["적용대손율(%)"]
DETAIL_DIFF_SIG = {
    "차이(계산-회사계상)": ("거래처별 회사계상 대손충당금(원)", SIG_THRESHOLD_ABS),
    "경과일수차이(계산-회사계산)": ("회사계산 경과일수(선택)", SIG_THRESHOLD_DAYS_ABS),
}


def save_results(df: pd.DataFrame, output_path: str, company: str, target_fy: str, 결산일: date,
                  labels: list, rate_table: dict, basis: dict, listed: str) -> None:
    wb = openpyxl.Workbook()

    header_fill = PatternFill("solid", fgColor="4472C4")
    header_font = Font(bold=True, color="FFFFFF")
    subtotal_fill = PatternFill("solid", fgColor="D9E1F2")
    total_fill = PatternFill("solid", fgColor="9DC3E6")
    sig_fill = PatternFill("solid", fgColor="FFFF00")
    bold = Font(bold=True)
    thin = Side(style="thin", color="B7B7B7")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    header_row = 3

    # ── 1. 매출채권명세 ──────────────────────────────────────────────────
    ws = wb.active
    ws.title = "매출채권명세"
    ws.cell(row=1, column=1,
            value=f"매출채권 연령분석/대손충당금 명세서 (회사: {company}, 회계연도: {target_fy}, "
                  f"상장구분: {listed}, 결산기준일: {결산일})").font = Font(bold=True, size=13)

    headers = list(df.columns) if not df.empty else list(DETAIL_MONEY_COLS)
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=header_row, column=i, value=h)
        c.fill = header_fill
        c.font = header_font
        c.alignment = center
        c.border = border
        ws.column_dimensions[get_column_letter(i)].width = 30 if h == "비고" else 16
    ws.freeze_panes = f"A{header_row + 1}"

    r = header_row + 1
    if df.empty:
        ws.cell(row=r, column=1, value="(입력된 매출채권 데이터 없음)")
    else:
        totals = {c: 0.0 for c in DETAIL_MONEY_COLS}
        for cat, gdf in df.groupby("평가구분", sort=False):
            for _, row in gdf.iterrows():
                for i, h in enumerate(headers, start=1):
                    val = row[h]
                    val = None if (val is None or (isinstance(val, float) and pd.isna(val))) else val
                    cell = ws.cell(row=r, column=i, value=val)
                    cell.border = border
                    if h in DETAIL_DATE_COLS and val is not None:
                        cell.number_format = "yyyy-mm-dd"
                    if h in DETAIL_MONEY_COLS and val is not None:
                        cell.number_format = "#,##0"
                    if h in DETAIL_DAY_COLS and val is not None:
                        cell.number_format = "0"
                    if h in DETAIL_PCT_COLS and val is not None:
                        cell.number_format = "0.00"
                    if h in DETAIL_DIFF_SIG and val is not None:
                        base_col, threshold = DETAIL_DIFF_SIG[h]
                        if _is_significant(val, row.get(base_col), threshold):
                            cell.fill = sig_fill
                for c in DETAIL_MONEY_COLS:
                    v = row.get(c)
                    if v is not None and not (isinstance(v, float) and pd.isna(v)):
                        totals[c] += v
                r += 1

            ws.cell(row=r, column=1, value=f"[{cat} 소계]").font = bold
            for i, h in enumerate(headers, start=1):
                cell = ws.cell(row=r, column=i)
                cell.fill = subtotal_fill
                cell.border = border
                if h in DETAIL_MONEY_COLS:
                    cell.value = totals[h]
                    cell.number_format = "#,##0"
                    cell.font = bold
            r += 1
            totals = {c: 0.0 for c in DETAIL_MONEY_COLS}

    # ── 2. 연령별대손율검증 ──────────────────────────────────────────────
    ws_bucket = wb.create_sheet("연령별대손율검증")
    bucket_rows = build_bucket_summary(df, labels, rate_table)
    ws_bucket.cell(row=1, column=1, value="연령구간별 대손율 적용 검증").font = Font(bold=True, size=13)

    bucket_headers = list(bucket_rows[0].keys()) if bucket_rows else []
    for i, h in enumerate(bucket_headers, start=1):
        c = ws_bucket.cell(row=header_row, column=i, value=h)
        c.fill = header_fill
        c.font = header_font
        c.alignment = center
        c.border = border
        ws_bucket.column_dimensions[get_column_letter(i)].width = 30 if h == "비고" else 20
    ws_bucket.freeze_panes = f"A{header_row + 1}"

    money_cols_bucket = {"순채권액(원)", "대손충당금(계산,원)"}
    pct_cols_bucket = {"회사설정 대손율(%)", "최근 실제대손율(참고,%)", "대손율차이(회사설정-실제,%p)"}
    r = header_row + 1
    for row in bucket_rows:
        for i, h in enumerate(bucket_headers, start=1):
            val = row[h]
            cell = ws_bucket.cell(row=r, column=i, value=val)
            cell.border = border
            if h in money_cols_bucket and val is not None:
                cell.number_format = "#,##0"
            if h in pct_cols_bucket and val is not None:
                cell.number_format = "0.00"
            if h == "비고" and val:
                cell.fill = sig_fill
        r += 1

    rate_warnings = check_rate_label_mismatch(labels, rate_table)
    if rate_warnings:
        r += 1
        for w in rate_warnings:
            ws_bucket.cell(row=r, column=1, value=w).font = Font(color="C00000", bold=True)
            r += 1

    # ── 3. 요약 ──────────────────────────────────────────────────────────
    ws_sum = wb.create_sheet("요약")
    ws_sum.cell(row=1, column=1,
                value=f"매출채권 대손충당금 검증 요약 (회사: {company}, 회계연도: {target_fy}, "
                      f"상장구분: {listed})").font = Font(bold=True, size=13)

    r = 3
    ws_sum.cell(row=r, column=1, value="[평가구분별 요약]").font = bold
    r += 1
    cat_rows = build_category_summary(df)
    cat_headers = list(cat_rows[0].keys()) if cat_rows else []
    for i, h in enumerate(cat_headers, start=1):
        c = ws_sum.cell(row=r, column=i, value=h)
        c.fill = header_fill
        c.font = header_font
        c.alignment = center
        c.border = border
        ws_sum.column_dimensions[get_column_letter(i)].width = 34 if h in ("비고", "평가구분") else 20
    r += 1
    for row in cat_rows:
        for i, h in enumerate(cat_headers, start=1):
            val = row[h]
            cell = ws_sum.cell(row=r, column=i, value=val)
            cell.border = border
            if h in ("순채권액(원)", "대손충당금(계산,원)") and val is not None:
                cell.number_format = "#,##0"
        r += 1

    overall = build_overall_summary(df, basis)
    r += 2
    ws_sum.cell(row=r, column=1, value="[전체 대사]").font = bold
    r += 1
    for label in ["총순채권액(집합+개별,원)", "총대손충당금(계산,원)", "당기말(회사계상,원)",
                  "당기말차이(계산-회사계상,원)"]:
        ws_sum.cell(row=r, column=1, value=label).border = border
        cell = ws_sum.cell(row=r, column=2, value=overall[label])
        cell.border = border
        if overall[label] is not None:
            cell.number_format = "#,##0"
        if label == "당기말차이(계산-회사계상,원)" and _is_significant(overall[label], overall["당기말(회사계상,원)"]):
            cell.fill = sig_fill
        r += 1

    r += 1
    ws_sum.cell(row=r, column=1, value="[설정률 전기 대비 비교]").font = bold
    r += 1
    for label in ["설정률(당기,재계산)", "설정률(전기,입력값기준)", "설정률차이(당기-전기,%p)"]:
        ws_sum.cell(row=r, column=1, value=label).border = border
        val = overall[label]
        cell = ws_sum.cell(row=r, column=2,
                            value=(val * 100.0 if val is not None and label != "설정률차이(당기-전기,%p)" else val))
        cell.border = border
        if val is not None:
            cell.number_format = "0.00"
        if label == "설정률차이(당기-전기,%p)" and overall["설정률_유의변동"]:
            cell.fill = sig_fill
            ws_sum.cell(row=r, column=3,
                        value="⚠ 전기 대비 설정률이 1%p 이상 변동 — 사유 확인 필요").font = Font(color="C00000")
        r += 1

    tie_out = overall["tie_out"]
    if tie_out:
        r += 1
        ws_sum.cell(row=r, column=1, value="[대손충당금 T계정 검증(tie-out)]").font = bold
        r += 1
        for label in ["전기말(회사계상)", "당기 전입액(입력)", "당기 환입액(입력)", "당기 직접상각액(입력)",
                      "계산상 당기말", "당기말(회사계상)", "차이(계산상당기말-회사계상)"]:
            ws_sum.cell(row=r, column=1, value=label).border = border
            cell = ws_sum.cell(row=r, column=2, value=tie_out[label])
            cell.border = border
            cell.number_format = "#,##0"
            if label == "차이(계산상당기말-회사계상)" and _is_significant(tie_out[label], tie_out["당기말(회사계상)"]):
                cell.fill = sig_fill
            r += 1

    r += 2
    ws_sum.cell(row=r, column=1, value="[상장구분별 안내]").font = bold
    r += 1
    if listed == "상장":
        fl = forward_looking_applied(basis)
        note = ("Forward-looking(미래전망정보) 조정 반영: " + ("Y — 근거 문서화 상태 확인 필요" if fl == "Y" else
                "N 또는 미입력 — K-IFRS9상 반영 필요 여부 및 근거를 회사에 확인 요망"))
        ws_sum.cell(row=r, column=1, value=note).font = Font(color="C00000" if fl != "Y" else "000000")
    else:
        ws_sum.cell(row=r, column=1,
                    value="비상장사 — '연령별대손율검증' 표의 '최근 실제대손율' 대비 회사설정 대손율 괴리(과소설정 "
                          "가능성) 경고를 우선 확인하세요.")
    r += 2

    ws_sum.cell(row=r, column=1, value="[참고: 계산 공식]").font = bold
    r += 1
    for text, is_bold in FORMULA_NOTE_LINES:
        cell = ws_sum.cell(row=r, column=1, value=text)
        if is_bold:
            cell.font = bold
        r += 1

    ws_sum.column_dimensions["A"].width = 46
    ws_sum.column_dimensions["B"].width = 20
    ws_sum.column_dimensions["C"].width = 50

    os.makedirs(os.path.dirname(output_path), exist_ok=True)
    wb.save(output_path)


# ── 메인 ─────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description="매출채권 연령분석/대손충당금 설정 검증앱")
    parser.add_argument("company", nargs="?", default=None, help="처리할 회사명 (생략 시 파일 자동 탐색)")
    parser.add_argument("--file", default=None, help="처리할 특정 입력 파일명 (input_data/ 기준)")
    parser.add_argument("--fiscal-month", type=int, default=12, help="결산월 (기본 12월). 예: 6월 결산법인이면 6")
    parser.add_argument("--fiscal-year", default=None, help="검증 대상 회계연도 (예: 2026). 생략 시 입력파일명의 fy 뒤 숫자 사용")
    parser.add_argument("--interim-month", type=int, default=None,
                         help="반기 등 중간결산 검토월 (예: 6). 당기말 결산기준일만 해당 월말로 앞당긴다")
    args = parser.parse_args()

    input_path = _find_input_file(args.company, args.file)
    company = args.company or os.path.basename(input_path).split("_")[2]

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
    결산일 = _ym_to_end_date(fy_end_ym)

    print(f"[입력] {input_path}")
    interim_note = f", 중간결산월={args.interim_month}(반기 등)" if args.interim_month else ""
    print(f"[대상] 회사={company}, 회계연도={target_fy}, 결산월={args.fiscal_month}{interim_note}")
    print(f"[결산기준일] {결산일}")

    receivables = load_receivables(input_path)
    rate_table = load_rate_table(input_path)
    basis = load_basis(input_path)
    listed = listed_type(basis)
    thresholds = parse_bucket_thresholds(basis)
    labels = bucket_labels(thresholds)

    print(f"[상장구분] {listed}, [연령구간] {labels}")
    print(f"[매출채권 건수] {len(receivables)}건")

    df = build_detail_table(receivables, 결산일, thresholds, labels, rate_table)

    suffix = f"_interim{args.interim_month:02d}" if args.interim_month else ""
    output_path = os.path.join(OUTPUT_DIR, f"ar_allowance_schedule_{company}_{target_fy}{suffix}.xlsx")
    save_results(df, output_path, company, target_fy, 결산일, labels, rate_table, basis, listed)
    print(f"[완료] {output_path}")


if __name__ == "__main__":
    main()
