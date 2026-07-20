"""
감사 조서 자동 분류기 (Audit Evidence Auto-Classifier)
------------------------------------------------------
실행 방법: python main.py
의존성: PyQt6, pandas, openpyxl
설정 파일: main.py와 동일 폴더에 category_map.xlsx 배치 필요

탭 구성
  Tab 1 — 자동 분류  : category_map.xlsx 키워드 기반 파일 일괄 이동
  Tab 2 — 미분류 정리: 미분류 파일을 작업자가 직접 선택·이동
"""

import sys
import os
import shutil
from datetime import datetime

import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QTextBrowser, QProgressBar,
    QMessageBox, QFrame, QTabWidget, QListWidget, QListWidgetItem,
    QComboBox, QLineEdit, QRadioButton, QButtonGroup, QGroupBox,
    QSplitter,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont


# ══════════════════════════════════════════════
# 1. 백그라운드 분류 스레드 (Tab 1 전용)
# ══════════════════════════════════════════════
class ClassifierWorker(QThread):
    """
    UI 블로킹 없이 파일 분류를 처리하는 백그라운드 스레드.
    시그널로 메인 스레드(UI)에 진행 상황을 전달한다.
    """
    log_signal      = pyqtSignal(str)   # 로그 한 줄
    progress_signal = pyqtSignal(int)   # 진행률 0~100
    finished_signal = pyqtSignal(dict)  # 완료 통계 {total, matched, unmatched}
    error_signal    = pyqtSignal(str)   # 치명적 오류

    def __init__(self, source_folder: str, target_folder: str, keyword_map: dict):
        super().__init__()
        self.source_folder = source_folder
        self.target_folder = target_folder
        self.keyword_map   = keyword_map

    @staticmethod
    def _normalize(text: str) -> str:
        """공백·괄호 제거 + 소문자 변환으로 비교 정확도를 높인다.
        '(별도) 보고서', '별도(보고서)', '별도 보고서' 모두 '별도보고서'로 동일하게 처리."""
        for ch in (' ', '(', ')', '[', ']', '{', '}'):
            text = text.replace(ch, '')
        return text.lower()

    def _safe_dest(self, dest_folder: str, filename: str) -> str:
        """동일 파일명 존재 시 '_HHMMSS' 타임스탬프를 붙여 충돌 방지."""
        dest = os.path.join(dest_folder, filename)
        if os.path.exists(dest):
            name, ext = os.path.splitext(filename)
            dest = os.path.join(dest_folder, f"{name}_{datetime.now().strftime('%H%M%S')}{ext}")
            self.log_signal.emit(f"    ※ 중복 파일명 자동 변경: {filename} → {os.path.basename(dest)}")
        return dest

    def _find_category(self, filename: str) -> str | None:
        """파일명(정규화)에 등록 키워드가 포함되면 계정과목명 반환, 없으면 None."""
        normalized = self._normalize(filename)
        for kw, cat in self.keyword_map.items():
            if kw in normalized:
                return cat
        return None

    def run(self):
        try:
            all_files = [
                f for f in os.listdir(self.source_folder)
                if os.path.isfile(os.path.join(self.source_folder, f))
            ]
            total = len(all_files)
            if total == 0:
                self.log_signal.emit("⚠ 원본 폴더에 처리할 파일이 없습니다.")
                self.finished_signal.emit({'total': 0, 'matched': 0, 'unmatched': 0})
                return

            self.log_signal.emit(f"총 {total}개 파일 처리를 시작합니다...\n")
            matched = unmatched = 0

            for idx, filename in enumerate(all_files):
                src = os.path.join(self.source_folder, filename)
                cat = self._find_category(filename)

                if cat:
                    dest_dir = os.path.join(self.target_folder, cat)
                    os.makedirs(dest_dir, exist_ok=True)
                    shutil.move(src, self._safe_dest(dest_dir, filename))
                    self.log_signal.emit(f"  [분류됨]  {filename}  →  [{cat}]")
                    matched += 1
                else:
                    dest_dir = os.path.join(self.target_folder, '00_미분류(확인필요)')
                    os.makedirs(dest_dir, exist_ok=True)
                    shutil.move(src, self._safe_dest(dest_dir, filename))
                    self.log_signal.emit(f"  [미분류]  {filename}  →  [00_미분류(확인필요)]")
                    unmatched += 1

                self.progress_signal.emit(int((idx + 1) / total * 100))

            self.finished_signal.emit({'total': total, 'matched': matched, 'unmatched': unmatched})
        except Exception as exc:
            self.error_signal.emit(f"처리 중 오류 발생: {exc}")


# ══════════════════════════════════════════════
# 2. Tab 1 — 자동 분류 위젯
# ══════════════════════════════════════════════
class AutoClassifyTab(QWidget):
    """category_map.xlsx 기반 자동 분류 탭."""

    # 대상 폴더가 바뀔 때 Tab 2 에 알리기 위한 시그널
    target_folder_changed = pyqtSignal(str)
    # 분류 완료 시 Tab 2 새로고침 요청
    classification_done   = pyqtSignal()

    def __init__(self, keyword_map: dict, parent=None):
        super().__init__(parent)
        self.keyword_map   = keyword_map
        self.source_folder = ""
        self.target_folder = ""
        self.worker        = None
        self._build_ui()

    # ── UI 구성 ────────────────────────────────
    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(14)
        layout.setContentsMargins(20, 16, 20, 12)

        # 원본 폴더 행
        layout.addLayout(self._folder_row(
            "원본 폴더  (클라이언트 자료)",
            "원본 폴더 선택", self._select_source, 'source_lbl'
        ))
        # 대상 폴더 행
        layout.addLayout(self._folder_row(
            "대상 폴더  (감사조서 저장 위치)",
            "대상 폴더 선택", self._select_target, 'target_lbl'
        ))

        # 분류 시작 버튼
        self.start_btn = QPushButton("▶  분류 시작")
        self.start_btn.setFixedHeight(46)
        self.start_btn.setStyleSheet("""
            QPushButton {
                background:#1D4ED8; color:white;
                font-size:14px; font-weight:bold; border-radius:7px;
            }
            QPushButton:hover    { background:#1E40AF; }
            QPushButton:pressed  { background:#1E3A8A; }
            QPushButton:disabled { background:#9CA3AF; color:#E5E7EB; }
        """)
        self.start_btn.clicked.connect(self._start)
        layout.addWidget(self.start_btn)

        # 진행률 바
        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.progress.setFixedHeight(18)
        layout.addWidget(self.progress)

        # 로그 창
        lbl = QLabel("처리 로그")
        lbl.setFont(QFont("", 10, QFont.Weight.Bold))
        layout.addWidget(lbl)

        self.log_box = QTextBrowser()
        self.log_box.setStyleSheet(
            "background:#F9FAFB; font-family:Consolas,monospace; font-size:11px;")
        layout.addWidget(self.log_box, 1)

        self.status_lbl = QLabel("준비")
        self.status_lbl.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.status_lbl.setStyleSheet("color:#6B7280; font-size:10px;")
        layout.addWidget(self.status_lbl)

    def _folder_row(self, title, btn_text, slot, attr) -> QVBoxLayout:
        """(안내 라벨 + 선택 버튼 + 경로 라벨) 행을 공통 생성."""
        col = QVBoxLayout()
        col.addWidget(self._small_bold(title))
        row = QHBoxLayout()
        btn = QPushButton(f"📁  {btn_text}")
        btn.setFixedWidth(200); btn.setFixedHeight(34)
        btn.setStyleSheet("""
            QPushButton { background:#F3F4F6; border:1px solid #D1D5DB; border-radius:5px; font-size:12px; }
            QPushButton:hover { background:#E5E7EB; }
        """)
        btn.clicked.connect(slot)
        lbl = QLabel("선택된 폴더 없음")
        lbl.setStyleSheet("color:#9CA3AF; border:1px solid #E5E7EB; padding:4px 8px; border-radius:5px; background:#FAFAFA;")
        lbl.setWordWrap(True)
        setattr(self, attr, lbl)
        row.addWidget(btn); row.addWidget(lbl, 1)
        col.addLayout(row)
        return col

    @staticmethod
    def _small_bold(text) -> QLabel:
        lbl = QLabel(text)
        lbl.setStyleSheet("color:#374151; font-weight:bold; font-size:11px;")
        return lbl

    # ── 슬롯 ───────────────────────────────────
    def _select_source(self):
        f = QFileDialog.getExistingDirectory(self, "원본 폴더 선택")
        if f:
            self.source_folder = f
            self.source_lbl.setText(f)
            self.source_lbl.setStyleSheet(
                "color:#111827; border:1px solid #D1D5DB; padding:4px 8px; border-radius:5px; background:#FAFAFA;")

    def _select_target(self):
        f = QFileDialog.getExistingDirectory(self, "대상 폴더 선택")
        if f:
            self.target_folder = f
            self.target_lbl.setText(f)
            self.target_lbl.setStyleSheet(
                "color:#111827; border:1px solid #D1D5DB; padding:4px 8px; border-radius:5px; background:#FAFAFA;")
            # Tab 2 에 대상 폴더 전달
            self.target_folder_changed.emit(f)

    def _start(self):
        if not self.source_folder:
            QMessageBox.warning(self, "입력 오류", "원본 폴더를 선택해주세요."); return
        if not self.target_folder:
            QMessageBox.warning(self, "입력 오류", "대상 폴더를 선택해주세요."); return
        if not self.keyword_map:
            QMessageBox.warning(self, "설정 오류", "category_map.xlsx 가 로드되지 않았습니다."); return
        if self.source_folder == self.target_folder:
            QMessageBox.warning(self, "입력 오류", "원본과 대상 폴더가 동일합니다."); return

        self.start_btn.setEnabled(False)
        self.progress.setValue(0)
        self.log_box.clear()
        self._log(f"▶ 분류 시작: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        self._log(f"  원본: {self.source_folder}")
        self._log(f"  대상: {self.target_folder}\n")

        self.worker = ClassifierWorker(self.source_folder, self.target_folder, self.keyword_map)
        self.worker.log_signal.connect(self._log)
        self.worker.progress_signal.connect(self.progress.setValue)
        self.worker.finished_signal.connect(self._on_finished)
        self.worker.error_signal.connect(self._on_error)
        self.worker.start()

    def _on_finished(self, s):
        self._log(f"\n{'─'*52}")
        self._log(f"✅ 분류 완료!  전체 {s['total']}개 | 분류됨 {s['matched']}개 | 미분류 {s['unmatched']}개")
        self._log(f"   완료 시각: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        self.status_lbl.setText(f"완료 — 총 {s['total']}개 (분류 {s['matched']} / 미분류 {s['unmatched']})")
        self.start_btn.setEnabled(True)
        self.classification_done.emit()   # Tab 2 새로고침 요청

    def _on_error(self, msg):
        self._log(f"\n❌ {msg}")
        self.status_lbl.setText("오류 발생")
        self.start_btn.setEnabled(True)
        QMessageBox.critical(self, "오류", msg)

    def _log(self, msg):
        self.log_box.append(msg)
        sb = self.log_box.verticalScrollBar()
        sb.setValue(sb.maximum())


# ══════════════════════════════════════════════
# 3. Tab 2 — 미분류 정리 위젯
# ══════════════════════════════════════════════
class ManualMoverTab(QWidget):
    """
    '00_미분류(확인필요)' 폴더의 파일을 작업자가 직접 선택하여
    기존 폴더 또는 새로 만들 폴더로 이동하는 탭.
    """

    UNCLASSIFIED_DIR = '00_미분류(확인필요)'

    def __init__(self, parent=None):
        super().__init__(parent)
        self.target_folder      = ""
        self.unclassified_path  = ""
        self._build_ui()

    # ── UI 구성 ────────────────────────────────
    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)
        layout.setContentsMargins(20, 16, 20, 12)

        # 안내 헤더
        self.folder_lbl = QLabel("대상 폴더: (자동 분류 탭에서 폴더를 선택하면 자동으로 연동됩니다)")
        self.folder_lbl.setStyleSheet("color:#6B7280; font-size:11px;")
        self.folder_lbl.setWordWrap(True)
        layout.addWidget(self.folder_lbl)

        line = QFrame(); line.setFrameShape(QFrame.Shape.HLine)
        line.setStyleSheet("color:#E5E7EB;")
        layout.addWidget(line)

        # ── 수평 스플리터: 파일 목록(좌) / 이동 옵션(우) ──
        splitter = QSplitter(Qt.Orientation.Horizontal)

        # 왼쪽: 파일 목록
        left = QWidget()
        lv = QVBoxLayout(left)
        lv.setContentsMargins(0, 0, 8, 0)

        top_row = QHBoxLayout()
        top_row.addWidget(self._bold_label("미분류 파일 목록"))
        self.file_count_lbl = QLabel("(0개)")
        self.file_count_lbl.setStyleSheet("color:#6B7280; font-size:11px;")
        top_row.addWidget(self.file_count_lbl)
        top_row.addStretch()
        refresh_btn = QPushButton("↻ 새로고침")
        refresh_btn.setFixedHeight(28)
        refresh_btn.setStyleSheet("font-size:11px; padding: 0 10px;")
        refresh_btn.clicked.connect(self.refresh_all)
        top_row.addWidget(refresh_btn)
        lv.addLayout(top_row)

        hint = QLabel("Ctrl+클릭 또는 Shift+클릭으로 복수 선택")
        hint.setStyleSheet("color:#9CA3AF; font-size:10px;")
        lv.addWidget(hint)

        self.file_list = QListWidget()
        self.file_list.setSelectionMode(QListWidget.SelectionMode.ExtendedSelection)
        self.file_list.setStyleSheet("font-size:12px;")
        lv.addWidget(self.file_list, 1)

        splitter.addWidget(left)

        # 오른쪽: 이동 옵션
        right = QWidget()
        rv = QVBoxLayout(right)
        rv.setContentsMargins(8, 0, 0, 0)

        rv.addWidget(self._bold_label("이동 대상 선택"))

        # ① 기존 폴더로 이동
        group = QGroupBox()
        group.setStyleSheet("QGroupBox { border:1px solid #E5E7EB; border-radius:6px; padding:10px; }")
        gv = QVBoxLayout(group)

        self.radio_existing = QRadioButton("기존 폴더로 이동")
        self.radio_existing.setChecked(True)
        self.radio_existing.setStyleSheet("font-weight:bold;")
        gv.addWidget(self.radio_existing)

        folder_row = QHBoxLayout()
        self.combo_folders = QComboBox()
        self.combo_folders.setFixedHeight(32)
        self.combo_folders.setStyleSheet("font-size:12px;")
        folder_row.addWidget(self.combo_folders, 1)

        refresh_folders_btn = QPushButton("↻")
        refresh_folders_btn.setFixedSize(32, 32)
        refresh_folders_btn.setToolTip("폴더 목록 새로고침")
        refresh_folders_btn.clicked.connect(self._refresh_folders)
        folder_row.addWidget(refresh_folders_btn)
        gv.addLayout(folder_row)

        # 구분선
        sep = QFrame(); sep.setFrameShape(QFrame.Shape.HLine)
        sep.setStyleSheet("color:#E5E7EB; margin:4px 0;")
        gv.addWidget(sep)

        # ② 새 폴더 만들기
        self.radio_new = QRadioButton("새 폴더 만들기")
        self.radio_new.setStyleSheet("font-weight:bold;")
        gv.addWidget(self.radio_new)

        self.new_folder_input = QLineEdit()
        self.new_folder_input.setPlaceholderText("새 폴더 이름을 입력하세요...")
        self.new_folder_input.setFixedHeight(32)
        self.new_folder_input.setEnabled(False)
        self.new_folder_input.setStyleSheet("font-size:12px; padding:2px 6px;")
        gv.addWidget(self.new_folder_input)

        rv.addWidget(group)

        # 라디오 버튼 연동
        self.btn_group = QButtonGroup()
        self.btn_group.addButton(self.radio_existing)
        self.btn_group.addButton(self.radio_new)
        self.radio_existing.toggled.connect(self._toggle_mode)

        rv.addStretch()

        # 이동 버튼
        self.move_btn = QPushButton("▶  선택 파일 이동")
        self.move_btn.setFixedHeight(44)
        self.move_btn.setStyleSheet("""
            QPushButton {
                background:#059669; color:white;
                font-size:13px; font-weight:bold; border-radius:7px;
            }
            QPushButton:hover   { background:#047857; }
            QPushButton:pressed { background:#065F46; }
            QPushButton:disabled{ background:#9CA3AF; }
        """)
        self.move_btn.clicked.connect(self._move_files)
        rv.addWidget(self.move_btn)

        splitter.addWidget(right)
        splitter.setSizes([420, 280])
        layout.addWidget(splitter, 1)

        # 로그 창 (하단 미니)
        self.log_box = QTextBrowser()
        self.log_box.setMaximumHeight(110)
        self.log_box.setStyleSheet(
            "background:#F0FDF4; font-family:Consolas,monospace; font-size:11px;")
        layout.addWidget(self.log_box)

    # ── 외부에서 호출: 폴더 세팅 ──────────────
    def set_target_folder(self, folder: str):
        """Tab 1 에서 대상 폴더가 설정될 때 호출된다."""
        self.target_folder     = folder
        self.unclassified_path = os.path.join(folder, self.UNCLASSIFIED_DIR)
        self.folder_lbl.setText(f"대상 폴더: {folder}")
        self.refresh_all()

    def refresh_all(self):
        """파일 목록 + 폴더 콤보박스를 동시에 새로고침한다."""
        self._refresh_files()
        self._refresh_folders()

    # ── 내부 새로고침 ──────────────────────────
    def _refresh_files(self):
        """미분류 폴더 내 파일 목록을 갱신한다."""
        self.file_list.clear()
        if not self.unclassified_path or not os.path.exists(self.unclassified_path):
            self.file_count_lbl.setText("(폴더 없음)")
            return
        files = sorted(
            f for f in os.listdir(self.unclassified_path)
            if os.path.isfile(os.path.join(self.unclassified_path, f))
        )
        for f in files:
            self.file_list.addItem(QListWidgetItem(f))
        self.file_count_lbl.setText(f"({len(files)}개)")

    def _refresh_folders(self):
        """대상 폴더 하위 폴더 목록(미분류 제외)을 콤보박스에 갱신한다."""
        self.combo_folders.clear()
        if not self.target_folder or not os.path.exists(self.target_folder):
            return
        folders = sorted(
            d for d in os.listdir(self.target_folder)
            if os.path.isdir(os.path.join(self.target_folder, d))
            and d != self.UNCLASSIFIED_DIR
        )
        self.combo_folders.addItems(folders)

    # ── 라디오 버튼 토글 ───────────────────────
    def _toggle_mode(self, existing_checked: bool):
        """기존 폴더 ↔ 새 폴더 입력 전환."""
        self.combo_folders.setEnabled(existing_checked)
        self.new_folder_input.setEnabled(not existing_checked)
        if not existing_checked:
            self.new_folder_input.setFocus()

    # ── 이동 실행 ──────────────────────────────
    def _move_files(self):
        """선택된 파일을 지정 폴더로 이동한다."""
        if not self.target_folder:
            QMessageBox.warning(self, "폴더 미설정",
                "자동 분류 탭에서 대상 폴더를 먼저 선택해주세요."); return

        selected = self.file_list.selectedItems()
        if not selected:
            QMessageBox.warning(self, "선택 없음", "이동할 파일을 선택해주세요."); return

        # 이동 대상 폴더 결정
        if self.radio_new.isChecked():
            new_name = self.new_folder_input.text().strip()
            if not new_name:
                QMessageBox.warning(self, "입력 오류", "새 폴더 이름을 입력해주세요."); return
            # 폴더 이름에 사용 불가 문자 체크
            invalid = set(r'\/:*?"<>|')
            if any(c in invalid for c in new_name):
                QMessageBox.warning(self, "이름 오류",
                    f"폴더 이름에 사용할 수 없는 문자가 포함되어 있습니다.\n( \\ / : * ? \" < > | )"); return
            dest_dir = os.path.join(self.target_folder, new_name)
        else:
            if self.combo_folders.count() == 0:
                QMessageBox.warning(self, "폴더 없음",
                    "이동 가능한 폴더가 없습니다.\n자동 분류를 먼저 실행하거나 새 폴더를 만들어주세요."); return
            dest_dir = os.path.join(self.target_folder, self.combo_folders.currentText())

        # 폴더 생성 확인
        try:
            os.makedirs(dest_dir, exist_ok=True)
        except Exception as exc:
            QMessageBox.critical(self, "폴더 생성 실패", str(exc)); return

        folder_name = os.path.basename(dest_dir)
        moved = errors = 0

        for item in selected:
            filename = item.text()
            src = os.path.join(self.unclassified_path, filename)
            dst = self._safe_dest(dest_dir, filename)
            try:
                shutil.move(src, dst)
                self._log(f"  이동: {filename}  →  [{folder_name}]")
                moved += 1
            except Exception as exc:
                self._log(f"  ❌ 오류: {filename}  —  {exc}")
                errors += 1

        summary = f"✅ {moved}개 이동 완료"
        if errors:
            summary += f"  (❌ {errors}개 오류)"
        self._log(f"\n{summary}\n")

        # 목록 새로고침
        self.refresh_all()

        # 새 폴더 입력란 초기화
        if self.radio_new.isChecked():
            self.new_folder_input.clear()

    # ── 중복 파일명 처리 ───────────────────────
    @staticmethod
    def _safe_dest(folder: str, filename: str) -> str:
        """동일 파일명 존재 시 '_HHMMSS' 접미사 추가."""
        path = os.path.join(folder, filename)
        if os.path.exists(path):
            name, ext = os.path.splitext(filename)
            path = os.path.join(folder, f"{name}_{datetime.now().strftime('%H%M%S')}{ext}")
        return path

    # ── 로그 ───────────────────────────────────
    def _log(self, msg: str):
        self.log_box.append(msg)
        sb = self.log_box.verticalScrollBar()
        sb.setValue(sb.maximum())

    # ── 유틸 ───────────────────────────────────
    @staticmethod
    def _bold_label(text: str) -> QLabel:
        lbl = QLabel(text)
        lbl.setFont(QFont("", 10, QFont.Weight.Bold))
        return lbl


# ══════════════════════════════════════════════
# 4. 메인 윈도우
# ══════════════════════════════════════════════
class AuditClassifierApp(QMainWindow):
    """감사 조서 자동 분류기 메인 윈도우 (탭 구조)."""

    def __init__(self, company=None, base=None):
        super().__init__()
        self._company = company
        self.keyword_map = {}
        self.setWindowTitle("감사 조서 자동 분류기")
        self.setMinimumSize(800, 620)

        self._load_category_map()   # 먼저 매핑 로드
        self._build_ui()            # 탭 UI 구성
        if company:
            import os as _os
            _project_root = _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__)))
            if base:
                _company_dir = _os.path.join(_project_root, base, company)
            else:
                _company_dir = _os.path.join(_project_root, company)
            if _os.path.isdir(_company_dir) and hasattr(self, 'tab_auto'):
                self.tab_auto.target_folder = _company_dir
                self.tab_auto.target_lbl.setText(_company_dir)
                self.tab_auto.target_lbl.setStyleSheet(
                    'color:#111827; border:1px solid #D1D5DB; padding:4px 8px; border-radius:5px; background:#FAFAFA;')
                self.tab_auto.target_folder_changed.emit(_company_dir)

    # ── 매핑 엑셀 로드 ──────────────────────────
    def _load_category_map(self):
        """
        스크립트(또는 실행 파일)와 동일 디렉터리의 category_map.xlsx를 읽어
        keyword_map 딕셔너리를 구성한다.

        엑셀 형식:
          폴더명         | 키워드
          현금및현금성자산 | 은행,잔액증명,cash,bank
        """
        base_dir = (os.path.dirname(sys.executable) if getattr(sys, 'frozen', False)
                    else os.path.dirname(os.path.abspath(__file__)))
        path = os.path.join(base_dir, 'category_map.xlsx')

        self._pending_logs = []   # UI 빌드 전 로그를 임시 저장

        if not os.path.exists(path):
            self._pending_logs.append(f"⚠ category_map.xlsx 를 찾을 수 없습니다.")
            self._pending_logs.append(f"  기대 경로: {path}")
            self._pending_logs.append("  main.py 와 같은 폴더에 파일을 넣고 재시작해주세요.\n")
            return

        try:
            df = pd.read_excel(path, dtype=str)
            if not {'폴더명', '키워드'}.issubset(df.columns):
                self._pending_logs.append(f"❌ 엑셀 컬럼 오류: '폴더명', '키워드' 컬럼이 필요합니다.")
                return

            kw_count = 0
            for _, row in df.iterrows():
                cat = str(row['폴더명']).strip()
                kws = str(row['키워드']).strip()
                if not kws or kws.lower() == 'nan' or not cat or cat.lower() == 'nan':
                    continue
                for kw in kws.split(','):
                    nkw = kw.strip().replace(' ', '').lower()
                    if nkw:
                        self.keyword_map[nkw] = cat
                        kw_count += 1

            self._pending_logs.append(f"✅ category_map.xlsx 로드 완료")
            self._pending_logs.append(f"   폴더명 {len(df)}개 · 키워드 {kw_count}개 등록\n")
        except Exception as exc:
            self._pending_logs.append(f"❌ category_map.xlsx 읽기 실패: {exc}")

    # ── UI 구성 ────────────────────────────────
    def _build_ui(self):
        root = QWidget()
        self.setCentralWidget(root)
        vbox = QVBoxLayout(root)
        vbox.setContentsMargins(0, 0, 0, 0)
        vbox.setSpacing(0)

        # 탭 위젯
        tabs = QTabWidget()
        tabs.setStyleSheet("""
            QTabBar::tab { padding: 8px 20px; font-size: 12px; }
            QTabBar::tab:selected { font-weight: bold; }
        """)

        # Tab 1
        self.tab_auto = AutoClassifyTab(self.keyword_map)
        tabs.addTab(self.tab_auto, "🗂  자동 분류")

        # Tab 2
        self.tab_manual = ManualMoverTab()
        tabs.addTab(self.tab_manual, "✏️  미분류 정리")

        # Tab 1 → Tab 2 연동
        self.tab_auto.target_folder_changed.connect(self.tab_manual.set_target_folder)
        self.tab_auto.classification_done.connect(self.tab_manual.refresh_all)

        vbox.addWidget(tabs)

        # 매핑 로드 결과를 Tab 1 로그에 출력
        for msg in self._pending_logs:
            self.tab_auto._log(msg)

        # 매핑 없으면 시작 버튼 비활성화
        if not self.keyword_map:
            self.tab_auto.start_btn.setEnabled(False)


# ══════════════════════════════════════════════
# 5. 진입점
# ══════════════════════════════════════════════
def main():
    import argparse
    _parser = argparse.ArgumentParser(add_help=False)
    _parser.add_argument('--company', default=None, help='회사 폴더명 (대상 폴더 자동 설정)')
    _parser.add_argument('--base', default=None, help='회사 폴더의 상위 기준 디렉터리 (예: journal_analyzer)')
    _args, _ = _parser.parse_known_args()
    app = QApplication(sys.argv)
    app.setStyle('Fusion')
    win = AuditClassifierApp(company=_args.company, base=_args.base)
    win.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
