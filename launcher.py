#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
감사 자동화 런처 — 도구 선택 및 실행 GUI
"""

import sys
import os
import re
from pathlib import Path
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QListWidget, QListWidgetItem, QLabel, QComboBox, QPushButton,
    QTextEdit, QGroupBox, QSplitter, QFrame, QSizePolicy, QLineEdit,
    QCheckBox,
)
from PyQt6.QtCore import Qt, QProcess, QProcessEnvironment
from PyQt6.QtGui import QFont, QColor, QTextCursor

ROOT = Path(__file__).parent

# 결산월 콤보박스 인덱스 → 결산월 매핑 (addItems 순서와 반드시 일치시킬 것)
_FY_MONTH_BY_INDEX = {0: 12, 1: 6, 2: 9}


# ──────────────────────────────────────────────────────────────────────────────
# 회사 목록 자동 감지
# ──────────────────────────────────────────────────────────────────────────────

def detect_js_companies() -> list[str]:
    """루트 폴더 중 task_list_*.xlsx가 있는 폴더를 회사로 간주"""
    companies = []
    for p in sorted(ROOT.iterdir()):
        if p.is_dir() and not p.name.startswith((".", "_")):
            if any(p.glob("task_list_*.xlsx")):
                companies.append(p.name)
    return companies


def detect_journal_companies() -> list[str]:
    """journal_analyzer/ 하위 폴더를 회사로 간주"""
    ja_dir = ROOT / "journal_analyzer"
    if not ja_dir.exists():
        return []
    return sorted(
        p.name for p in ja_dir.iterdir()
        if p.is_dir() and not p.name.startswith((".", "_", "__"))
    )


def detect_interest_companies() -> list[str]:
    """interest_analyzer/ 하위 폴더를 회사로 간주"""
    ia_dir = ROOT / "interest_analyzer"
    if not ia_dir.exists():
        return []
    return sorted(
        p.name for p in ia_dir.iterdir()
        if p.is_dir() and not p.name.startswith((".", "_", "__"))
    )


def detect_lease_companies() -> list[str]:
    """lease_analyzer/input_data/ 의 파일명에서 회사명 추출"""
    in_dir = ROOT / 'lease_analyzer' / 'input_data'
    if not in_dir.exists():
        return []
    _pat = re.compile(r'^lease_(.+)_information_', re.IGNORECASE)
    companies = set()
    for f in in_dir.iterdir():
        if f.is_file() and f.suffix.lower() == '.xlsx' and not f.name.startswith('~$'):
            m = _pat.match(f.name)
            if m:
                companies.add(m.group(1))
    return sorted(companies)


def detect_lease_input_files() -> list[tuple[str, str]]:
    """lease_analyzer/input_data/ 내 파일별 (표시명, 파일명) 목록 반환"""
    in_dir = ROOT / 'lease_analyzer' / 'input_data'
    if not in_dir.exists():
        return []
    _pat = re.compile(r'^lease_(.+)_information_(.+)\.xlsx$', re.IGNORECASE)
    files = []
    for f in sorted(in_dir.iterdir()):
        if f.is_file() and f.suffix.lower() == '.xlsx' and not f.name.startswith('~$'):
            m = _pat.match(f.name)
            if m:
                company, year = m.group(1), m.group(2)
                files.append((f'{company}  [{year}]', f.name))
    return files


def detect_dep_input_files() -> list[tuple[str, str]]:
    """depreciation_analyzer/input_data/ 내 파일별 (표시명, 파일명) 목록 반환"""
    in_dir = ROOT / 'depreciation_analyzer' / 'input_data'
    if not in_dir.exists():
        return []
    _pat = re.compile(r'^depreciation_(.+)_information_(.+)\.xlsx$', re.IGNORECASE)
    files = []
    for f in sorted(in_dir.iterdir()):
        if f.is_file() and f.suffix.lower() == '.xlsx' and not f.name.startswith('~$') and 'template' not in f.name.lower():
            m = _pat.match(f.name)
            if m:
                company, year = m.group(1), m.group(2)
                files.append((f'{company}  [{year}]', f.name))
    return files


def detect_sev_input_files() -> list[tuple[str, str]]:
    """severance_analyzer/input_data/ 내 파일별 (표시명, 파일명) 목록 반환"""
    in_dir = ROOT / 'severance_analyzer' / 'input_data'
    if not in_dir.exists():
        return []
    _pat = re.compile(r'^severance_(.+)_information_(.+)\.xlsx$', re.IGNORECASE)
    files = []
    for f in sorted(in_dir.iterdir()):
        if f.is_file() and f.suffix.lower() == '.xlsx' and not f.name.startswith('~$') and 'template' not in f.name.lower():
            m = _pat.match(f.name)
            if m:
                company, year = m.group(1), m.group(2)
                files.append((f'{company}  [{year}]', f.name))
    return files

# ──────────────────────────────────────────────────────────────────────────────
# 도구 정의
# ──────────────────────────────────────────────────────────────────────────────

TOOLS = [
    {
        "category": "감사자료 extractor",
        "name": "Playwright 자동화 (run.js)",
        "desc": "task_list 기반 웹 자동화 (계정별원장 다운로드, 분개장 AI 분석 등)",
        "cmd": ["node", "run.js"],
        "cwd": str(ROOT),
        "company": "js",        # detect_js_companies()
        "extra": None,
    },
    {
        "category": "분개장분석_파이썬",
        "name": "분개장분석_파이썬 (journal_analyzer)",
        "desc": "분개장 데이터를 분석번호 선택 실행 (벤포드·거래처비교·심층분석 등)\n비워두면 task_list에서 Y 표시된 전체 분석 실행",
        "cmd": ["python", str(ROOT / "journal_analyzer" / "main_analyzer.py")],
        "cwd": str(ROOT),
        "company": "journal",   # detect_journal_companies()
        "extra": "task_filter",
    },
    {
        "category": "Data_Injector",
        "name": "데이터 주입 - JS회사 (data_injector)",
        "desc": "mapping_list 기반으로 감사조서 엑셀에 데이터를 자동 주입 (루트 회사)",
        "cmd": ["python", str(ROOT / "report" / "data_injector.py")],
        "cwd": str(ROOT / "report"),
        "company": "js",
        "extra": None,
    },
    {
        "category": "Data_Injector",
        "name": "데이터 주입 - journal회사 (data_injector)",
        "desc": "mapping_list 기반으로 감사조서 엑셀에 데이터를 자동 주입 (journal_analyzer 회사)",
        "cmd": ["python", str(ROOT / "report" / "data_injector.py"), "--base", "journal_analyzer"],
        "cwd": str(ROOT / "report"),
        "company": "journal",
        "extra": None,
    },
    {
        "category": "리스 분석",
        "name": "리스 스케줄 생성 (lease_schedule)",
        "desc": "⚠ 실행 전 lease_analyzer/input_data/ 폴더에 lease_{회사명}_information_{연도}.xlsx 파일을 먼저 업로드하세요.\nK-IFRS 1116 리스 회계처리 — 계약별 요약 + 월별 상각 스케줄 생성",
        "cmd": ["python", str(ROOT / "lease_analyzer" / "lease_schedule.py")],
        "cwd": str(ROOT / "lease_analyzer"),
        "company": "lease_file",
        "extra": None,
    },
    {
        "category": "리스 분석",
        "name": "리스 완전성 검토 - JS회사 (lease_filter)",
        "desc": "K-IFRS 1116 리스 완전성 검토 — 회사 미선택 시 input_data 직접분석, 선택 시 results 연계",
        "cmd": ["python", str(ROOT / "lease_analyzer" / "lease_filter.py")],
        "cwd": str(ROOT / "lease_analyzer"),
        "company": "optional_js",
        "extra": "optional_company",
    },
    {
        "category": "리스 분석",
        "name": "리스 완전성 검토 - journal회사 (lease_filter)",
        "desc": "K-IFRS 1116 리스 완전성 검토 — journal_analyzer 회사의 results 폴더 연계",
        "cmd": ["python", str(ROOT / "lease_analyzer" / "lease_filter.py"), "--base", "journal_analyzer"],
        "cwd": str(ROOT / "lease_analyzer"),
        "company": "journal",
        "extra": None,
    },
    {
        "category": "감가상각 검증",
        "name": "감가상각비 검증 (depreciation_schedule)",
        "desc": "⚠ 실행 전 depreciation_analyzer/input_data/ 폴더에 depreciation_{회사명}_information_fy{연도}.xlsx 파일을 먼저 업로드하세요 (build_template.py로 표준 템플릿 생성 가능).\n유형자산·투자부동산 감가상각비 재계산(정액법/정률법) — 계정과목별 소계가 있는 고정자산명세서 생성, 회사계상액과 자동 대사",
        "cmd": ["python", str(ROOT / "depreciation_analyzer" / "depreciation_schedule.py")],
        "cwd": str(ROOT / "depreciation_analyzer"),
        "company": "dep_file",
        "extra": None,
    },
    {
        "category": "퇴직급여충당부채 검증",
        "name": "퇴직급여충당부채 검증 (severance_schedule)",
        "desc": "⚠ 실행 전 severance_analyzer/input_data/ 폴더에 severance_{회사명}_information_fy{연도}.xlsx 파일을 먼저 업로드하세요 (build_template.py로 표준 템플릿 생성 가능, 파일 안에 '당기정보'/'전기정보' 두 시트 포함).\n일반기업회계기준 퇴직금 추계액 방식 — 인원별 재직일수·급여로 당기말 잔액 재계산, 회사계상액과 자동 대사, 신규입사자/퇴사자 명단 자동 산출",
        "cmd": ["python", str(ROOT / "severance_analyzer" / "severance_schedule.py")],
        "cwd": str(ROOT / "severance_analyzer"),
        "company": "sev_file",
        "extra": None,
    },
    {
        "category": "시트 분리",
        "name": "엑셀 시트 분리 (sheet_splitter)",
        "desc": "시트명 기준(sheet) 또는 컬럼값 기준(col)으로 엑셀 파일을 분리",
        "cmd": ["python", str(ROOT / "sheet_splitter" / "split_sheets.py")],
        "cwd": str(ROOT / "sheet_splitter"),
        "company": None,
        "extra": "sheet_mode",  # 모드 선택
    },
    {
        "category": "파일 분류",
        "name": "감사조서 파일 분류 - JS회사 (file_classifier)",
        "desc": "오딧로비 업로드를 위한 파일 분류 — 키워드 기반 감사 조서 자동 분류 GUI (루트 회사)",
        "cmd": ["python", str(ROOT / "file_classifier" / "main.py")],
        "cwd": str(ROOT / "file_classifier"),
        "company": "js",
        "extra": "company_flag",
    },
    {
        "category": "파일 분류",
        "name": "감사조서 파일 분류 - journal회사 (file_classifier)",
        "desc": "오딧로비 업로드를 위한 파일 분류 — 키워드 기반 감사 조서 자동 분류 GUI (journal_analyzer 회사)",
        "cmd": ["python", str(ROOT / "file_classifier" / "main.py"), "--base", "journal_analyzer"],
        "cwd": str(ROOT / "file_classifier"),
        "company": "journal",
        "extra": "company_flag",
    },
    {
        "category": "적정성 분석",
        "name": "이자비용 적정성 분석 (interest_expense_analysis)",
        "desc": "차입금 잔액 기반 일별 이자 계산 → 실제 이자비용과 비교 검증 (interest_analyzer/<회사>/input)",
        "cmd": ["python", str(ROOT / "interest_analyzer" / "interest_expense_analysis.py")],
        "cwd": str(ROOT / "interest_analyzer"),
        "company": "interest",
        "extra": None,
    },
    {
        "category": "주석 검증",
        "name": "감사주석 검증 - JS회사 (note_verifier)",
        "desc": "정산표/DSD와 감사조서 주석을 시트별 비교 검증 — 블록1 자동채움 + 차이 색상 표시 (루트 회사)",
        "cmd": ["python", str(ROOT / "note_verifier" / "note_verifier.py")],
        "cwd": str(ROOT),
        "company": "js",
        "extra": "company_flag",
    },
    {
        "category": "주석 검증",
        "name": "감사주석 검증 - journal회사 (note_verifier)",
        "desc": "정산표/DSD와 감사조서 주석을 시트별 비교 검증 — 블록1 자동채움 + 차이 색상 표시 (journal_analyzer 회사)",
        "cmd": ["python", str(ROOT / "note_verifier" / "note_verifier.py"), "--base", "journal_analyzer"],
        "cwd": str(ROOT),
        "company": "journal",
        "extra": "company_flag",
    },
]


# ──────────────────────────────────────────────────────────────────────────────
# 메인 윈도우
# ──────────────────────────────────────────────────────────────────────────────

class Launcher(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("감사 자동화 런처")
        self.resize(900, 640)
        self._process: QProcess | None = None

        self._js_companies = detect_js_companies()
        self._journal_companies = detect_journal_companies()
        self._interest_companies = detect_interest_companies()
        self._lease_companies = detect_lease_companies()
        self._lease_input_files: list[tuple[str, str]] = detect_lease_input_files()
        self._dep_input_files: list[tuple[str, str]] = detect_dep_input_files()
        self._sev_input_files: list[tuple[str, str]] = detect_sev_input_files()

        self._build_ui()
        self._tool_list.setCurrentRow(0)

    # ── UI 구성 ──────────────────────────────────────────────────────────────

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root_layout = QVBoxLayout(central)
        root_layout.setContentsMargins(12, 12, 12, 12)
        root_layout.setSpacing(8)

        # 제목
        title = QLabel("감사 자동화 런처")
        title.setFont(QFont("맑은 고딕", 14, QFont.Weight.Bold))
        root_layout.addWidget(title)

        sep = QFrame()
        sep.setFrameShape(QFrame.Shape.HLine)
        sep.setFrameShadow(QFrame.Shadow.Sunken)
        root_layout.addWidget(sep)

        # 상단: 도구 목록 + 설정 패널
        splitter = QSplitter(Qt.Orientation.Horizontal)
        root_layout.addWidget(splitter, stretch=1)

        # 왼쪽: 도구 목록
        left = QGroupBox("도구 목록")
        left_layout = QVBoxLayout(left)
        self._tool_list = QListWidget()
        self._tool_list.setFont(QFont("맑은 고딕", 10))
        for t in TOOLS:
            item = QListWidgetItem(f"[{t['category']}]  {t['name'].split('(')[0].strip()}")
            self._tool_list.addItem(item)
        self._tool_list.currentRowChanged.connect(self._on_tool_selected)
        left_layout.addWidget(self._tool_list)
        splitter.addWidget(left)

        # 오른쪽: 설정 패널
        right = QGroupBox("실행 설정")
        right_layout = QVBoxLayout(right)
        right_layout.setSpacing(10)

        self._lbl_name = QLabel("")
        self._lbl_name.setFont(QFont("맑은 고딕", 11, QFont.Weight.Bold))
        self._lbl_name.setWordWrap(True)
        right_layout.addWidget(self._lbl_name)

        self._lbl_desc = QLabel("")
        self._lbl_desc.setFont(QFont("맑은 고딕", 9))
        self._lbl_desc.setWordWrap(True)
        self._lbl_desc.setStyleSheet("color: #555;")
        right_layout.addWidget(self._lbl_desc)

        sep2 = QFrame()
        sep2.setFrameShape(QFrame.Shape.HLine)
        sep2.setFrameShadow(QFrame.Shadow.Sunken)
        right_layout.addWidget(sep2)

        # 회사 선택
        self._company_row = QWidget()
        co_layout = QHBoxLayout(self._company_row)
        co_layout.setContentsMargins(0, 0, 0, 0)
        co_layout.addWidget(QLabel("회사 선택:"))
        self._company_combo = QComboBox()
        self._company_combo.setFont(QFont("맑은 고딕", 10))
        self._company_combo.setMinimumWidth(160)
        co_layout.addWidget(self._company_combo)
        co_layout.addStretch()
        right_layout.addWidget(self._company_row)

        # 시트 필터 (JS 자동화 전용: --sheet 옵션)
        self._sheet_row = QWidget()
        sheet_layout = QHBoxLayout(self._sheet_row)
        sheet_layout.setContentsMargins(0, 0, 0, 0)
        sheet_layout.addWidget(QLabel("task_list 시트명(부분일치):"))
        self._sheet_input = QLineEdit()
        self._sheet_input.setFont(QFont("맑은 고딕", 10))
        self._sheet_input.setPlaceholderText("예: 벤포드법칙  또는  벤포드법칙, 전기  (쉼표로 다중 지정 / 비워두면 전체 실행)")
        self._sheet_input.setMinimumWidth(220)
        sheet_layout.addWidget(self._sheet_input)
        sheet_layout.addStretch()
        right_layout.addWidget(self._sheet_row)

        # 분析번호 필터 (journal_analyzer 전용: --task 옵션)
        self._task_row = QWidget()
        task_layout = QHBoxLayout(self._task_row)
        task_layout.setContentsMargins(0, 0, 0, 0)
        task_layout.addWidget(QLabel("분析번호 선택:"))
        self._task_input = QLineEdit()
        self._task_input.setFont(QFont("맑은 고딕", 10))
        self._task_input.setPlaceholderText("예: 14  또는  3 8 14 22  (공백 구분 / 비워두면 Y 전체 실행)")
        self._task_input.setMinimumWidth(220)
        task_layout.addWidget(self._task_input)
        task_layout.addStretch()
        right_layout.addWidget(self._task_row)

        # 시트 분리 모드
        self._mode_row = QWidget()
        mode_layout = QHBoxLayout(self._mode_row)
        mode_layout.setContentsMargins(0, 0, 0, 0)
        mode_layout.addWidget(QLabel("분리 모드:"))
        self._mode_combo = QComboBox()
        self._mode_combo.setFont(QFont("맑은 고딕", 10))
        self._mode_combo.addItems(["sheet  (시트명 기준)", "col  (컬럼값 기준)"])
        mode_layout.addWidget(self._mode_combo)
        mode_layout.addStretch()
        right_layout.addWidget(self._mode_row)

        # 결산월 선택 (리스 스케줄 전용)
        self._fy_row = QWidget()
        fy_layout = QHBoxLayout(self._fy_row)
        fy_layout.setContentsMargins(0, 0, 0, 0)
        fy_layout.addWidget(QLabel("결산월:"))
        self._fy_combo = QComboBox()
        self._fy_combo.setFont(QFont("맑은 고딕", 10))
        self._fy_combo.addItems(["12월 결산 (1월~12월)", "6월 결산 (반기, 1월~6월)", "9월 결산 (1월~9월)"])
        fy_layout.addWidget(self._fy_combo)
        fy_layout.addStretch()
        right_layout.addWidget(self._fy_row)

        # 반기 등 중간결산 검토 (리스 스케줄·감가상각 검증 공용) — 12월 결산 법인의 상반기(1~6월) 검토처럼
        # 회계연도는 그대로 두고 계산기간만 특정월까지로 앞당길 때 사용 (결산월 자체를 바꾸는 위 콤보와는 별개)
        self._interim_row = QWidget()
        interim_layout = QHBoxLayout(self._interim_row)
        interim_layout.setContentsMargins(0, 0, 0, 0)
        self._interim_check = QCheckBox("반기(1~6월)만 계산 — 12월 결산 법인의 상반기 검토용 (기초잔액은 전기말 그대로, 당기는 1~6월만)")
        self._interim_check.setFont(QFont("맑은 고딕", 9))
        interim_layout.addWidget(self._interim_check)
        interim_layout.addStretch()
        right_layout.addWidget(self._interim_row)

        right_layout.addStretch()

        # 실행 / 중지 버튼
        btn_row = QHBoxLayout()
        self._btn_run = QPushButton("▶  실행")
        self._btn_run.setFont(QFont("맑은 고딕", 11, QFont.Weight.Bold))
        self._btn_run.setMinimumHeight(38)
        self._btn_run.setStyleSheet(
            "QPushButton { background:#2563EB; color:white; border-radius:6px; }"
            "QPushButton:hover { background:#1D4ED8; }"
            "QPushButton:disabled { background:#9CA3AF; }"
        )
        self._btn_run.clicked.connect(self._run)

        self._btn_stop = QPushButton("■  중지")
        self._btn_stop.setFont(QFont("맑은 고딕", 11))
        self._btn_stop.setMinimumHeight(38)
        self._btn_stop.setEnabled(False)
        self._btn_stop.setStyleSheet(
            "QPushButton { background:#DC2626; color:white; border-radius:6px; }"
            "QPushButton:hover { background:#B91C1C; }"
            "QPushButton:disabled { background:#9CA3AF; }"
        )
        self._btn_stop.clicked.connect(self._stop)

        btn_row.addWidget(self._btn_run)
        btn_row.addWidget(self._btn_stop)
        right_layout.addLayout(btn_row)

        splitter.addWidget(right)
        splitter.setSizes([320, 560])

        # 하단: 로그 패널
        log_group = QGroupBox("실행 로그")
        log_layout = QVBoxLayout(log_group)
        self._log = QTextEdit()
        self._log.setReadOnly(True)
        self._log.setTextInteractionFlags(
            Qt.TextInteractionFlag.TextSelectableByMouse |
            Qt.TextInteractionFlag.TextSelectableByKeyboard
        )
        self._log.setFont(QFont("Consolas", 9))
        self._log.setStyleSheet("background:#1E1E1E; color:#D4D4D4;")
        log_layout.addWidget(self._log)

        log_btn_row = QHBoxLayout()
        log_btn_row.addStretch()
        btn_copy_log = QPushButton("로그 복사")
        btn_copy_log.setMaximumWidth(100)
        def _copy_log():
            try:
                cb = QApplication.clipboard()
                cb.setText(self._log.toPlainText())
            except Exception:
                pass
        btn_copy_log.clicked.connect(_copy_log)
        btn_save_log = QPushButton("로그 저장")
        btn_save_log.setMaximumWidth(100)
        def _save_log():
            import datetime, os
            log_dir = ROOT / "docs" / "logs"
            log_dir.mkdir(parents=True, exist_ok=True)
            fname = log_dir / f"launcher_log_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
            fname.write_text(self._log.toPlainText(), encoding="utf-8")
            self._log.append(f"[로그 저장 완료] {fname}")
        btn_save_log.clicked.connect(_save_log)
        btn_clear = QPushButton("로그 지우기")
        btn_clear.setMaximumWidth(100)
        btn_clear.clicked.connect(self._log.clear)
        log_btn_row.addWidget(btn_copy_log)
        log_btn_row.addWidget(btn_save_log)
        log_btn_row.addWidget(btn_clear)
        log_layout.addLayout(log_btn_row)
        root_layout.addWidget(log_group, stretch=1)

    # ── 도구 선택 시 패널 업데이트 ──────────────────────────────────────────

    def _on_tool_selected(self, idx: int):
        if idx < 0 or idx >= len(TOOLS):
            return
        t = TOOLS[idx]
        self._lbl_name.setText(t["name"])
        self._lbl_desc.setText(t["desc"])

        # 회사 선택 콤보
        self._company_combo.blockSignals(True)
        self._company_combo.clear()
        if t["company"] == "js":
            self._company_row.setVisible(True)
            self._company_combo.addItems(self._js_companies)
        elif t["company"] == "journal":
            self._company_row.setVisible(True)
            self._company_combo.addItems(self._journal_companies)
        elif t["company"] == "optional_js":
            self._company_row.setVisible(True)
            self._company_combo.addItem("(없음 — input_data 직접분석)")
            self._company_combo.addItems(self._js_companies)
        elif t["company"] == "interest":
            self._company_row.setVisible(True)
            self._company_combo.addItems(self._interest_companies)
        elif t["company"] == "lease":
            self._company_row.setVisible(True)
            self._company_combo.addItems(self._lease_companies)
        elif t["company"] == "lease_file":
            self._company_row.setVisible(True)
            self._lease_input_files = detect_lease_input_files()
            self._company_combo.addItems([d for d, _ in self._lease_input_files])
        elif t["company"] == "dep_file":
            self._company_row.setVisible(True)
            self._dep_input_files = detect_dep_input_files()
            self._company_combo.addItems([d for d, _ in self._dep_input_files])
        elif t["company"] == "sev_file":
            self._company_row.setVisible(True)
            self._sev_input_files = detect_sev_input_files()
            self._company_combo.addItems([d for d, _ in self._sev_input_files])
        else:
            self._company_row.setVisible(False)
        self._company_combo.blockSignals(False)

        # 시트 필터 (JS 자동화 전용)
        self._sheet_row.setVisible(t["company"] == "js")
        if t["company"] == "js":
            self._sheet_input.clear()

        # 분析번호 필터 (journal_analyzer 전용)
        self._task_row.setVisible(t.get("extra") == "task_filter")
        if t.get("extra") == "task_filter":
            self._task_input.clear()

        # 모드 선택 (sheet_splitter 전용)
        self._mode_row.setVisible(t["extra"] == "sheet_mode")

        # 결산월 선택 (리스 스케줄·감가상각 검증·퇴충 검증 전용)
        self._fy_row.setVisible(t["company"] in ("lease_file", "dep_file", "sev_file"))

        # 반기 등 중간결산 검토 (리스 스케줄·감가상각 검증·퇴충 검증 전용)
        self._interim_row.setVisible(t["company"] in ("lease_file", "dep_file", "sev_file"))
        if t["company"] not in ("lease_file", "dep_file", "sev_file"):
            self._interim_check.setChecked(False)

    # ── 실행 ─────────────────────────────────────────────────────────────────

    def _run(self):
        idx = self._tool_list.currentRow()
        if idx < 0:
            return
        t = TOOLS[idx]

        cmd = list(t["cmd"])

        # 회사명 인자 추가
        if t["company"] in ("js", "journal", "interest", "lease"):
            company = self._company_combo.currentText().strip()
            if not company:
                self._log_line("⚠  회사를 선택하세요.", "#FBBF24")
                return
            if t.get("extra") == "company_flag":
                cmd += ["--company", company]
            else:
                cmd.append(company)
        elif t["company"] == "optional_js":
            selected = self._company_combo.currentText()
            if not selected.startswith("(없음"):
                cmd += ["--company", selected]
        elif t["company"] == "lease_file":
            file_idx = self._company_combo.currentIndex()
            if file_idx < 0 or not self._lease_input_files:
                self._log_line("⚠  처리할 파일을 선택하세요.", "#FBBF24")
                return
            _, fname = self._lease_input_files[file_idx]
            cmd += ["--file", fname]
            fy_month = _FY_MONTH_BY_INDEX.get(self._fy_combo.currentIndex(), 12)
            if self._interim_check.isChecked():
                # 반기(1~6월) 검토: 결산월 6월 기준 스냅샷은 유지하되 당기 집계는 1~6월로 제한
                cmd += ["--fiscal-month", "6", "--interim"]
            elif fy_month in (6, 9):
                # 결산월 콤보에서 "6월/9월 결산" 선택 — 모두 12월 결산 법인의 중간결산 검토용
                # (같은 캘린더 연도의 1월~fy_month월 누적)이므로 동일하게 --interim으로 처리
                cmd += ["--fiscal-month", str(fy_month), "--interim"]
            else:
                cmd += ["--fiscal-month", str(fy_month)]
        elif t["company"] == "dep_file":
            file_idx = self._company_combo.currentIndex()
            if file_idx < 0 or not self._dep_input_files:
                self._log_line("⚠  처리할 파일을 선택하세요.", "#FBBF24")
                return
            _, fname = self._dep_input_files[file_idx]
            cmd += ["--file", fname]
            fy_month = _FY_MONTH_BY_INDEX.get(self._fy_combo.currentIndex(), 12)
            if self._interim_check.isChecked():
                # 반기(1~6월) 검토: 결산월은 12월 그대로, 계산기간만 6월까지로 앞당김
                cmd += ["--fiscal-month", "12", "--interim-month", "6"]
            elif fy_month in (6, 9):
                # 결산월 콤보에서 "6월/9월 결산" 선택 — 모두 12월 결산 법인의 중간결산 검토용
                # (결산월은 12월 그대로, 계산기간만 해당월까지로 앞당김 → 파일명에 _interim06/_interim09 부여)
                cmd += ["--fiscal-month", "12", "--interim-month", str(fy_month)]
            else:
                cmd += ["--fiscal-month", str(fy_month)]
        elif t["company"] == "sev_file":
            file_idx = self._company_combo.currentIndex()
            if file_idx < 0 or not self._sev_input_files:
                self._log_line("⚠  처리할 파일을 선택하세요.", "#FBBF24")
                return
            _, fname = self._sev_input_files[file_idx]
            cmd += ["--file", fname]
            fy_month = _FY_MONTH_BY_INDEX.get(self._fy_combo.currentIndex(), 12)
            if self._interim_check.isChecked():
                # 반기(1~6월) 검토: 결산월은 12월 그대로, 당기말 결산기준일만 6월말로 앞당김
                cmd += ["--fiscal-month", "12", "--interim-month", "6"]
            elif fy_month in (6, 9):
                # 결산월 콤보에서 "6월/9월 결산" 선택 — 모두 12월 결산 법인의 중간결산 검토용
                cmd += ["--fiscal-month", "12", "--interim-month", str(fy_month)]
            else:
                cmd += ["--fiscal-month", str(fy_month)]

        # 시트 필터 인자 추가 (JS 자동화 전용)
        if t["company"] == "js":
            sheet_filter = self._sheet_input.text().strip()
            if sheet_filter:
                cmd += ["--sheet", sheet_filter]

        # 분析번호 필터 인자 추가 (journal_analyzer 전용)
        if t.get("extra") == "task_filter":
            task_text = self._task_input.text().strip()
            if task_text:
                task_nums = [s for s in task_text.replace(",", " ").split() if s.isdigit()]
                if task_nums:
                    cmd += ["--task"] + task_nums

        # 시트 분리 모드 인자 추가
        if t["extra"] == "sheet_mode":
            mode = "sheet" if self._mode_combo.currentIndex() == 0 else "col"
            cmd += ["--mode", mode]

        self._log_line(f"\n▶ 실행: {' '.join(cmd)}", "#60A5FA")

        self._process = QProcess(self)
        env = QProcessEnvironment.systemEnvironment()
        env.insert("PYTHONIOENCODING", "utf-8")
        env.insert("PYTHONUTF8", "1")
        self._process.setProcessEnvironment(env)
        self._process.setWorkingDirectory(t["cwd"])
        self._process.readyReadStandardOutput.connect(self._on_stdout)
        self._process.readyReadStandardError.connect(self._on_stderr)
        self._process.finished.connect(self._on_finished)

        self._btn_run.setEnabled(False)
        self._btn_stop.setEnabled(True)

        self._process.start(cmd[0], cmd[1:])
        if not self._process.waitForStarted(3000):
            self._log_line("✗ 프로세스를 시작할 수 없습니다.", "#F87171")
            self._reset_buttons()

    def _stop(self):
        if self._process and self._process.state() != QProcess.ProcessState.NotRunning:
            self._process.kill()
            self._log_line("■ 프로세스를 강제 종료했습니다.", "#FBBF24")

    # ── 프로세스 이벤트 ───────────────────────────────────────────────────────

    def _on_stdout(self):
        data = bytes(self._process.readAllStandardOutput()).decode("utf-8", errors="replace")
        self._log_line(data.rstrip(), "#D4D4D4")

    def _on_stderr(self):
        data = bytes(self._process.readAllStandardError()).decode("utf-8", errors="replace")
        self._log_line(data.rstrip(), "#F87171")

    def _on_finished(self, exit_code: int, _):
        color = "#4ADE80" if exit_code == 0 else "#F87171"
        self._log_line(f"{'✓' if exit_code == 0 else '✗'}  종료 (exit {exit_code})", color)
        self._reset_buttons()

    def _reset_buttons(self):
        self._btn_run.setEnabled(True)
        self._btn_stop.setEnabled(False)

    def _log_line(self, text: str, color: str = "#D4D4D4"):
        self._log.moveCursor(QTextCursor.MoveOperation.End)
        self._log.setTextColor(QColor(color))
        self._log.insertPlainText(text + "\n")
        self._log.moveCursor(QTextCursor.MoveOperation.End)


# ──────────────────────────────────────────────────────────────────────────────

def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    win = Launcher()
    win.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
