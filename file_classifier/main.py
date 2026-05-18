"""
감사 조서 자동 분류기 (Audit Evidence Auto-Classifier)
------------------------------------------------------
실행 방법: python main.py
의존성: PyQt6, pandas, openpyxl
설정 파일: main.py와 동일 폴더에 category_map.xlsx 배치 필요
"""

import sys
import os
import shutil
from datetime import datetime

import pandas as pd
from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QFileDialog, QTextBrowser, QProgressBar,
    QMessageBox, QFrame,
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont, QColor


# ──────────────────────────────────────────────
# 백그라운드 작업 스레드
# ──────────────────────────────────────────────
class ClassifierWorker(QThread):
    """
    UI가 멈추지 않도록 파일 분류 로직을 별도 스레드에서 실행한다.
    시그널을 통해 메인 스레드(UI)에 진행 상황을 전달한다.
    """

    log_signal      = pyqtSignal(str)   # 로그 한 줄 전달
    progress_signal = pyqtSignal(int)   # 진행률(0~100) 전달
    finished_signal = pyqtSignal(dict)  # 완료 통계 {total, matched, unmatched} 전달
    error_signal    = pyqtSignal(str)   # 치명적 오류 메시지 전달

    def __init__(self, source_folder: str, target_folder: str, keyword_map: dict):
        super().__init__()
        self.source_folder = source_folder  # 원본 폴더 경로 (클라이언트 자료)
        self.target_folder = target_folder  # 대상 폴더 경로 (감사조서 저장 위치)
        self.keyword_map   = keyword_map    # {정규화된_키워드: 계정과목명}

    # ── 정규화: 공백 제거 + 소문자 변환 ─────────────────────────────
    @staticmethod
    def _normalize(text: str) -> str:
        """비교 정확도 향상을 위해 공백 제거 및 영어 소문자 통일"""
        return text.replace(' ', '').lower()

    # ── 중복 파일명 처리 ─────────────────────────────────────────────
    def _safe_dest_path(self, dest_folder: str, filename: str) -> str:
        """
        대상 폴더에 동일한 파일명이 이미 존재하면
        파일명 끝에 '_HHMMSS' 타임스탬프를 붙여 충돌을 방지한다.
        원본 파일은 절대 삭제·덮어쓰기하지 않는다.
        """
        dest_path = os.path.join(dest_folder, filename)
        if os.path.exists(dest_path):
            name, ext = os.path.splitext(filename)
            timestamp = datetime.now().strftime('%H%M%S')
            new_name  = f"{name}_{timestamp}{ext}"
            dest_path = os.path.join(dest_folder, new_name)
            self.log_signal.emit(f"    ※ 중복 파일명 자동 변경: {filename} → {new_name}")
        return dest_path

    # ── 키워드 매칭 ──────────────────────────────────────────────────
    def _find_category(self, filename: str) -> str | None:
        """
        파일명(정규화)에 등록된 키워드가 포함되어 있으면
        해당 계정과목명을 반환한다. 없으면 None.
        """
        normalized = self._normalize(filename)
        for keyword, category in self.keyword_map.items():
            if keyword in normalized:
                return category
        return None

    # ── 메인 실행 로직 ───────────────────────────────────────────────
    def run(self):
        """스레드 진입점: 원본 폴더 파일을 순회하며 분류·이동한다."""
        try:
            # 원본 폴더 직접 하위 파일만 수집 (하위 폴더 내 파일 제외)
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

            matched_count   = 0
            unmatched_count = 0

            for idx, filename in enumerate(all_files):
                src_path = os.path.join(self.source_folder, filename)
                category = self._find_category(filename)

                if category:
                    # ── 매칭 성공: 계정과목 이름의 하위 폴더로 이동 ──
                    dest_folder = os.path.join(self.target_folder, category)
                    os.makedirs(dest_folder, exist_ok=True)
                    dest_path = self._safe_dest_path(dest_folder, filename)
                    shutil.move(src_path, dest_path)
                    self.log_signal.emit(f"  [분류됨]  {filename}  →  [{category}]")
                    matched_count += 1
                else:
                    # ── 매칭 실패: '00_미분류(확인필요)' 폴더로 이동 ──
                    unmatched_folder = os.path.join(self.target_folder, '00_미분류(확인필요)')
                    os.makedirs(unmatched_folder, exist_ok=True)
                    dest_path = self._safe_dest_path(unmatched_folder, filename)
                    shutil.move(src_path, dest_path)
                    self.log_signal.emit(f"  [미분류]  {filename}  →  [00_미분류(확인필요)]")
                    unmatched_count += 1

                # 진행률 시그널 (0 ~ 100)
                self.progress_signal.emit(int((idx + 1) / total * 100))

            self.finished_signal.emit({
                'total':     total,
                'matched':   matched_count,
                'unmatched': unmatched_count,
            })

        except Exception as exc:
            self.error_signal.emit(f"처리 중 오류 발생: {exc}")


# ──────────────────────────────────────────────
# 메인 윈도우
# ──────────────────────────────────────────────
class AuditClassifierApp(QMainWindow):
    """감사 조서 자동 분류기 메인 윈도우"""

    def __init__(self):
        super().__init__()
        self.source_folder = ""   # 선택된 원본 폴더 경로
        self.target_folder = ""   # 선택된 대상 폴더 경로
        self.keyword_map   = {}   # {정규화된_키워드: 계정과목명}
        self.worker        = None # 백그라운드 작업 스레드 참조

        self.setWindowTitle("감사 조서 자동 분류기")
        self.setMinimumSize(750, 600)

        self._build_ui()          # UI 컴포넌트 구성
        self._load_category_map() # 매핑 엑셀 로드

    # ── UI 구성 ──────────────────────────────────────────────────────
    def _build_ui(self):
        """위젯과 레이아웃을 조립하여 화면을 구성한다."""
        root = QWidget()
        self.setCentralWidget(root)
        layout = QVBoxLayout(root)
        layout.setSpacing(14)
        layout.setContentsMargins(24, 20, 24, 16)

        # ── 제목 ──
        title = QLabel("감사 조서 자동 분류기")
        f = QFont()
        f.setPointSize(17)
        f.setBold(True)
        title.setFont(f)
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(title)

        # ── 구분선 ──
        line = QFrame()
        line.setFrameShape(QFrame.Shape.HLine)
        line.setStyleSheet("color: #D1D5DB;")
        layout.addWidget(line)

        # ── 원본 폴더 선택 행 ──
        layout.addLayout(self._folder_row(
            label_text  = "원본 폴더  (클라이언트 자료)",
            btn_text    = "원본 폴더 선택",
            slot        = self._select_source,
            attr_label  = 'source_label',
        ))

        # ── 대상 폴더 선택 행 ──
        layout.addLayout(self._folder_row(
            label_text  = "대상 폴더  (감사조서 저장 위치)",
            btn_text    = "대상 폴더 선택",
            slot        = self._select_target,
            attr_label  = 'target_label',
        ))

        # ── 분류 시작 버튼 ──
        self.start_btn = QPushButton("▶  분류 시작")
        self.start_btn.setFixedHeight(46)
        self.start_btn.setStyleSheet("""
            QPushButton {
                background-color: #1D4ED8;
                color: white;
                font-size: 14px;
                font-weight: bold;
                border-radius: 7px;
            }
            QPushButton:hover    { background-color: #1E40AF; }
            QPushButton:pressed  { background-color: #1E3A8A; }
            QPushButton:disabled { background-color: #9CA3AF; color: #E5E7EB; }
        """)
        self.start_btn.clicked.connect(self._start)
        layout.addWidget(self.start_btn)

        # ── 진행률 바 ──
        self.progress = QProgressBar()
        self.progress.setValue(0)
        self.progress.setFixedHeight(18)
        self.progress.setTextVisible(True)
        layout.addWidget(self.progress)

        # ── 로그 창 ──
        log_title = QLabel("처리 로그")
        log_title.setFont(QFont("", 10, QFont.Weight.Bold))
        layout.addWidget(log_title)

        self.log_box = QTextBrowser()
        self.log_box.setStyleSheet(
            "background:#F9FAFB; font-family:Consolas,monospace; font-size:11px;"
        )
        layout.addWidget(self.log_box, 1)  # stretch=1 → 남은 공간 모두 사용

        # ── 하단 상태 라벨 ──
        self.status_lbl = QLabel("준비 완료")
        self.status_lbl.setAlignment(Qt.AlignmentFlag.AlignRight)
        self.status_lbl.setStyleSheet("color: #6B7280; font-size: 10px;")
        layout.addWidget(self.status_lbl)

    def _folder_row(self, label_text: str, btn_text: str, slot, attr_label: str) -> QHBoxLayout:
        """
        (안내 라벨) + (선택 버튼) + (경로 표시 라벨) 로 구성된
        폴더 선택 행 레이아웃을 공통으로 생성한다.
        """
        row = QVBoxLayout()
        info = QLabel(label_text)
        info.setStyleSheet("color: #374151; font-weight: bold; font-size: 11px;")
        row.addWidget(info)

        inner = QHBoxLayout()
        btn = QPushButton(f"📁  {btn_text}")
        btn.setFixedWidth(200)
        btn.setFixedHeight(34)
        btn.setStyleSheet("""
            QPushButton {
                background: #F3F4F6;
                border: 1px solid #D1D5DB;
                border-radius: 5px;
                font-size: 12px;
            }
            QPushButton:hover { background: #E5E7EB; }
        """)
        btn.clicked.connect(slot)

        path_lbl = QLabel("선택된 폴더 없음")
        path_lbl.setStyleSheet(
            "color:#9CA3AF; border:1px solid #E5E7EB;"
            "padding:4px 8px; border-radius:5px; background:#FAFAFA;"
        )
        path_lbl.setWordWrap(True)

        # 동적 속성 등록 (self.source_label / self.target_label)
        setattr(self, attr_label, path_lbl)

        inner.addWidget(btn)
        inner.addWidget(path_lbl, 1)
        row.addLayout(inner)
        return row

    # ── 매핑 엑셀 로드 ───────────────────────────────────────────────
    def _load_category_map(self):
        """
        스크립트(또는 실행 파일)와 동일한 디렉터리에서 category_map.xlsx를 읽어
        keyword_map 딕셔너리를 구성한다.

        엑셀 형식:
          계정과목   | 키워드
          현금및현금성자산 | 은행,잔액증명,cash,bank
          매출채권    | 채권,receivable,invoice
        """
        # PyInstaller 패키징 여부에 따라 기준 디렉터리 결정
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))

        excel_path = os.path.join(base_dir, 'category_map.xlsx')

        if not os.path.exists(excel_path):
            self._log(f"⚠ category_map.xlsx 를 찾을 수 없습니다.")
            self._log(f"  → 기대 경로: {excel_path}")
            self._log("  → 파일을 main.py 와 같은 폴더에 넣고 재시작해주세요.\n")
            self.status_lbl.setText("category_map.xlsx 없음 — 분류 불가")
            self.start_btn.setEnabled(False)
            return

        try:
            df = pd.read_excel(excel_path, dtype=str)

            # 필수 컬럼 존재 여부 확인
            required = {'계정과목', '키워드'}
            if not required.issubset(df.columns):
                missing = required - set(df.columns)
                self._log(f"❌ 엑셀 컬럼 오류: {missing} 컬럼이 없습니다.")
                self.start_btn.setEnabled(False)
                return

            self.keyword_map = {}
            kw_count = 0

            for _, row in df.iterrows():
                category = str(row['계정과목']).strip()
                raw_kws  = str(row['키워드']).strip()

                # 빈 셀 또는 NaN 스킵
                if not raw_kws or raw_kws.lower() == 'nan' or not category or category.lower() == 'nan':
                    continue

                # 쉼표 구분 키워드를 각각 정규화(공백제거+소문자)하여 등록
                for kw in raw_kws.split(','):
                    normalized = kw.strip().replace(' ', '').lower()
                    if normalized:
                        self.keyword_map[normalized] = category
                        kw_count += 1

            self._log(f"✅ category_map.xlsx 로드 완료")
            self._log(f"   계정과목 {len(df)}개 · 키워드 {kw_count}개 등록\n")
            self.status_lbl.setText(f"매핑 로드 완료 ({kw_count}개 키워드)")

        except Exception as exc:
            self._log(f"❌ category_map.xlsx 읽기 실패: {exc}")
            self.start_btn.setEnabled(False)

    # ── 폴더 선택 슬롯 ───────────────────────────────────────────────
    def _select_source(self):
        """원본 폴더 선택 다이얼로그"""
        folder = QFileDialog.getExistingDirectory(self, "원본 폴더 선택 (클라이언트 자료)")
        if folder:
            self.source_folder = folder
            self.source_label.setText(folder)
            self.source_label.setStyleSheet(
                "color:#111827; border:1px solid #D1D5DB;"
                "padding:4px 8px; border-radius:5px; background:#FAFAFA;"
            )

    def _select_target(self):
        """대상 폴더 선택 다이얼로그"""
        folder = QFileDialog.getExistingDirectory(self, "대상 폴더 선택 (감사조서 저장 위치)")
        if folder:
            self.target_folder = folder
            self.target_label.setText(folder)
            self.target_label.setStyleSheet(
                "color:#111827; border:1px solid #D1D5DB;"
                "padding:4px 8px; border-radius:5px; background:#FAFAFA;"
            )

    # ── 분류 시작 ────────────────────────────────────────────────────
    def _start(self):
        """입력값 검증 후 백그라운드 분류 작업을 시작한다."""
        # ── 사전 검증 ──
        if not self.source_folder:
            QMessageBox.warning(self, "입력 오류", "원본 폴더를 선택해주세요.")
            return
        if not self.target_folder:
            QMessageBox.warning(self, "입력 오류", "대상 폴더를 선택해주세요.")
            return
        if not self.keyword_map:
            QMessageBox.warning(self, "설정 오류",
                "category_map.xlsx 가 로드되지 않았습니다.\n"
                "파일을 확인한 후 프로그램을 재시작해주세요.")
            return
        if self.source_folder == self.target_folder:
            QMessageBox.warning(self, "입력 오류",
                "원본 폴더와 대상 폴더가 동일합니다.\n다른 폴더를 선택해주세요.")
            return

        # ── UI 비활성화(작업 중) ──
        self.start_btn.setEnabled(False)
        self.progress.setValue(0)
        self.log_box.clear()
        self._log(f"▶ 분류 시작: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        self._log(f"  원본: {self.source_folder}")
        self._log(f"  대상: {self.target_folder}\n")

        # ── 백그라운드 스레드 실행 ──
        self.worker = ClassifierWorker(self.source_folder, self.target_folder, self.keyword_map)
        self.worker.log_signal.connect(self._log)
        self.worker.progress_signal.connect(self.progress.setValue)
        self.worker.finished_signal.connect(self._on_finished)
        self.worker.error_signal.connect(self._on_error)
        self.worker.start()

    # ── 완료 콜백 ────────────────────────────────────────────────────
    def _on_finished(self, stats: dict):
        """분류 완료 후 결과를 로그에 요약하고 UI를 복원한다."""
        self._log(f"\n{'─' * 52}")
        self._log(f"✅ 분류 완료!")
        self._log(f"   전체 {stats['total']}개  |  "
                  f"분류됨 {stats['matched']}개  |  "
                  f"미분류 {stats['unmatched']}개")
        self._log(f"   완료 시각: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        self.status_lbl.setText(
            f"완료 — 총 {stats['total']}개 처리 "
            f"(분류 {stats['matched']} / 미분류 {stats['unmatched']})"
        )
        self.start_btn.setEnabled(True)

    # ── 오류 콜백 ────────────────────────────────────────────────────
    def _on_error(self, msg: str):
        """치명적 오류 발생 시 로그 출력 및 UI 복원"""
        self._log(f"\n❌ {msg}")
        self.status_lbl.setText("오류 발생 — 로그를 확인해주세요")
        self.start_btn.setEnabled(True)
        QMessageBox.critical(self, "오류", msg)

    # ── 로그 출력 헬퍼 ───────────────────────────────────────────────
    def _log(self, message: str):
        """로그 창에 한 줄을 추가하고 자동으로 최하단으로 스크롤한다."""
        self.log_box.append(message)
        sb = self.log_box.verticalScrollBar()
        sb.setValue(sb.maximum())


# ──────────────────────────────────────────────
# 진입점
# ──────────────────────────────────────────────
def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")           # 크로스 플랫폼 일관된 디자인
    window = AuditClassifierApp()
    window.show()
    sys.exit(app.exec())


if __name__ == '__main__':
    main()
