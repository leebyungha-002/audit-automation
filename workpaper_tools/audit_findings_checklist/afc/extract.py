"""원문 수집·텍스트 추출 (PDF / HWP / HWPX). 페이지 번호를 보존하고, 이미 처리한 파일은 해시로 건너뛴다."""
import hashlib
import re
from pathlib import Path

import pymupdf

from .db import Store, now

SOURCE_SUFFIXES = {".pdf", ".hwp", ".hwpx"}
PRIVATE_USE = re.compile(r"[-]")  # 한글 글머리표 등 사용자 정의 영역 문자


def file_sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


def detect_issuer(file_name: str, head_text: str, rules: list[dict]) -> str:
    for rule in rules:
        target = file_name if rule["where"] == "filename" else head_text
        if re.search(rule["pattern"], target):
            return rule["issuer"]
    return "미상"


def decide_mode(file_name: str, images_per_page: float, ecfg: dict) -> str:
    if any(s in file_name for s in ecfg["force_vision"]):
        return "vision"
    if any(s in file_name for s in ecfg["force_text"]):
        return "text"
    return "vision" if images_per_page >= ecfg["vision_min_images_per_page"] else "text"


def clean_page_text(text: str, noise: list[re.Pattern]) -> str:
    lines = [PRIVATE_USE.sub("", ln).rstrip() for ln in text.splitlines()]
    return "\n".join(ln for ln in lines if ln.strip() and not any(p.match(ln.strip()) for p in noise))


def render_page_png(pdf_path: Path, page_no: int, dpi: int) -> bytes:
    with pymupdf.open(pdf_path) as doc:
        return doc[page_no - 1].get_pixmap(dpi=dpi).tobytes("png")


class HwpReader:
    """한컴오피스 자동화로 HWP/HWPX의 페이지별 텍스트를 읽는다 (한글이 설치된 PC에서만 동작)."""

    def __init__(self):
        self._hwp = None

    def read_pages(self, path: Path) -> list[str]:
        if self._hwp is None:
            import win32com.client
            self._hwp = win32com.client.Dispatch("HWPFrame.HwpObject")
        fmt = "HWPX" if path.suffix.lower() == ".hwpx" else "HWP"
        if not self._hwp.Open(str(path), fmt, "forceopen:true"):
            raise RuntimeError("한글에서 파일을 열지 못함")
        return [self._hwp.GetPageText(i, "") for i in range(self._hwp.PageCount)]

    def close(self) -> None:
        if self._hwp is not None:
            self._hwp.Clear(1)  # 저장하지 않고 닫기
            self._hwp.Quit()
            self._hwp = None


def _read_pdf(path: Path) -> tuple[list[str], int]:
    with pymupdf.open(path) as doc:
        return [page.get_text() for page in doc], sum(len(page.get_images()) for page in doc)


def run_extract(cfg: dict, store: Store, log, force: bool = False, name_filter: str | None = None) -> None:
    ecfg = cfg["extract"]
    noise = [re.compile(p) for p in ecfg["noise_lines"]]
    files = sorted(p for p in cfg["paths"]["input_pdfs"].iterdir()
                   if p.suffix.lower() in SOURCE_SUFFIXES and not p.name.startswith("~"))
    files = [p for p in files if not any(s in p.name for s in ecfg["exclude_files"])]
    if name_filter:
        files = [p for p in files if name_filter in p.name]
    if not files:
        log.warning("처리할 파일이 없습니다: %s", cfg["paths"]["input_pdfs"])
        return

    hwp = HwpReader()
    try:
        for src in files:
            fhash = file_sha256(src)
            if not force and store.query("SELECT 1 FROM source_files WHERE file_hash=?", (fhash,)):
                log.info("건너뜀(이미 처리됨): %s", src.name)
                continue
            try:
                if src.suffix.lower() == ".pdf":
                    raw, n_images = _read_pdf(src)
                else:
                    raw, n_images = hwp.read_pages(src), 0
            except Exception as e:  # 손상/암호화 파일, 한글 미설치 등
                log.error("파일 읽기 실패: %s (%s)", src.name, e)
                store.log("extract", "ERROR", src.name, f"파일 읽기 실패: {e}")
                continue

            n_pages = len(raw)
            all_text = "".join(raw)
            space_ratio = all_text.count(" ") / max(len(all_text), 1)
            mode = decide_mode(src.name, n_images / max(n_pages, 1), ecfg) if src.suffix.lower() == ".pdf" else "text"
            year_match = re.search(r"(20\d{2})", src.name)
            issuer = detect_issuer(src.name, "".join(raw[:3]), cfg["issuer_rules"])

            store.conn.execute("DELETE FROM pages WHERE file_hash=?", (fhash,))
            ocr_pages = []
            for i, text in enumerate(raw, start=1):
                cleaned = clean_page_text(text, noise)
                n_chars = len(re.sub(r"\s+", "", cleaned))
                needs_ocr = int(n_chars < ecfg["ocr_min_chars"])
                if needs_ocr:
                    ocr_pages.append(i)
                store.upsert("pages", {"file_hash": fhash, "page_no": i, "text": cleaned,
                                       "n_chars": n_chars, "needs_ocr": needs_ocr})
            store.upsert("source_files", {
                "file_hash": fhash, "file_name": src.name, "issuer": issuer,
                "doc_year": int(year_match.group(1)) if year_match else None,
                "n_pages": n_pages, "extract_mode": mode, "processed_at": now(),
            })
            store.commit()

            log.info("추출 완료: %s | %d쪽 | 발행 %s | 모드 %s (공백비율 %.2f, 쪽당 이미지 %.1f)",
                     src.name, n_pages, issuer, mode, space_ratio, n_images / max(n_pages, 1))
            store.log("extract", "INFO", src.name, f"{n_pages}쪽 추출, 모드={mode}, 발행={issuer}")
            if mode == "vision":
                msg = "글리프 이미지 의심(숫자·문장부호 누락 가능) → 구조화 시 페이지 이미지 병행"
                log.warning("%s: %s", src.name, msg)
                store.log("extract", "WARNING", src.name, msg)
            elif space_ratio < ecfg["low_space_ratio"]:
                msg = f"띄어쓰기 소실(공백비율 {space_ratio:.2f}) — 발췌문은 AI가 띄어쓰기를 복원함"
                log.warning("%s: %s", src.name, msg)
                store.log("extract", "WARNING", src.name, msg)
            if ocr_pages:
                msg = f"텍스트 거의 없는 페이지(OCR 대상): {ocr_pages}"
                log.warning("%s: %s", src.name, msg)
                store.log("extract", "WARNING", src.name, msg)
    finally:
        hwp.close()
