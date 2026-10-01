"""감리 지적사항 ↔ 계정리스트 대사 체크리스트 생성기 CLI.

사용 예:
    python main.py run-all                 # 새 원문만 추출·분할하고 구조화 예상 비용을 보여줌
    python main.py run-all --yes --batch   # 위에 이어 구조화까지 실행 (배치: 50% 할인, 수 분~1시간)
    python main.py extract                 # 원문(PDF/HWP) 텍스트 추출
    python main.py segment                 # 지적사항 단위 분할
    python main.py structure --limit 5     # Claude API로 구조화 (처음 5건만)
    python main.py review-export           # 검토용 엑셀 내보내기 (output/review_날짜.xlsx)
    python main.py review-import           # 수정한 검토 엑셀을 DB에 재반영 (경로 생략 시 최신 파일)
    python main.py map --accounts <파일>   # 계정리스트를 표준 분류로 정규화하고 지적사항과 매칭
    python main.py report --company <회사> # 체크리스트 엑셀 출력 (output/checklist_회사_날짜.xlsx)
    python main.py status                  # 현황 요약
"""
import argparse
import sys
from pathlib import Path

from afc.config import load_config, load_env, setup_logging
from afc.db import Store
from afc.extract import run_extract
from afc.mapping import run_map
from afc.report import run_report
from afc.review import STATUSES, export_review, import_review
from afc.segment import run_segment
from afc.structure import estimate, run_structure


def show_status(cfg: dict, store: Store, log) -> None:
    print("\n[원문 파일]")
    for r in store.query(
        "SELECT f.file_name, f.issuer, f.n_pages, f.extract_mode, "
        "(SELECT COUNT(*) FROM segments s WHERE s.file_hash=f.file_hash) AS n_seg, "
        "(SELECT COUNT(*) FROM findings d JOIN segments s ON s.segment_id=d.segment_id "
        " WHERE s.file_hash=f.file_hash) AS n_find FROM source_files f ORDER BY f.file_name"
    ):
        print(f"  {r['file_name']} | {r['issuer']} | {r['n_pages']}쪽 | {r['extract_mode']} | "
              f"분할 {r['n_seg']}건 | 구조화 {r['n_find']}건")
    pending, cost = estimate(cfg, store, log)
    print(f"\n[구조화 대기] {pending}건 (예상 비용 약 ${cost:.2f}, 배치 사용 시 약 ${cost / 2:.2f})")
    for r in store.query("SELECT batch_id, n_requests, created_at FROM batches WHERE status='submitted'"):
        print(f"  진행 중인 배치: {r['batch_id']} ({r['n_requests']}건, {r['created_at']} 제출)")
    print("\n[검토 상태]")
    for r in store.query("SELECT review_status, COUNT(*) AS n FROM findings GROUP BY review_status"):
        print(f"  {r['review_status']}: {r['n']}건")
    print("\n[발췌 검증]")
    for r in store.query("SELECT excerpt_status, COUNT(*) AS n FROM findings GROUP BY excerpt_status"):
        print(f"  {r['excerpt_status']}: {r['n']}건")
    print("\n[계정 매칭]")
    for r in store.query("SELECT company, COUNT(*) AS n, SUM(category IS NULL) AS none FROM account_mappings GROUP BY company"):
        print(f"  {r['company']}: 계정 {r['n']}개 (매칭 없음 {r['none']}개)")
    print("\n[경고/오류 로그]")
    for r in store.query("SELECT level, COUNT(*) AS n FROM process_log WHERE level<>'INFO' GROUP BY level"):
        print(f"  {r['level']}: {r['n']}건")


def run_all(cfg: dict, store: Store, log, yes: bool, batch: bool, limit: int | None, wait: int) -> None:
    """폴더 일괄 처리. 이미 처리한 파일·사례는 건너뛰므로 새 원문을 넣고 다시 실행하면 증분만 처리된다."""
    run_extract(cfg, store, log)
    run_segment(cfg, store, log)
    pending, cost = estimate(cfg, store, log, batch=batch)
    open_batch = store.query("SELECT 1 FROM batches WHERE status='submitted'")
    if pending == 0 and not open_batch:
        log.info("새로 구조화할 지적사례가 없습니다.")
        return
    log.info("구조화 대기 %d건, 예상 비용 약 $%.2f%s", pending, cost, " (배치 50%% 할인 적용)" if batch else "")
    if not yes:
        log.info("구조화는 API 비용이 발생하므로 실행하지 않았습니다. 진행하려면 --yes를 붙여 다시 실행하세요 "
                 "(--batch를 함께 쓰면 반값).")
        return
    run_structure(cfg, store, log, limit=limit, batch=batch, wait_minutes=wait)
    log.info("다음 순서: review-export → (엑셀 검토) → review-import → map --accounts <파일> → report --company <회사>")


def main() -> int:
    parser = argparse.ArgumentParser(description="감리 지적사항 체크리스트 생성기")
    sub = parser.add_subparsers(dest="command", required=True)
    for name in ("extract", "segment", "structure"):
        p = sub.add_parser(name)
        p.add_argument("--force", action="store_true", help="이미 처리한 것도 다시 처리")
        p.add_argument("--file", help="파일명에 이 문자열이 포함된 것만 처리")
        if name == "segment":
            p.add_argument("--llm", action="store_true", help="규칙 대신 LLM 보조 분할 강제")
    p = sub.add_parser("run-all", help="새 원문 추출·분할 후 구조화까지 일괄 처리")
    p.add_argument("--yes", action="store_true", help="예상 비용 확인 없이 구조화까지 실행")
    for name in ("structure", "run-all"):
        p = sub.choices[name]
        p.add_argument("--limit", type=int, help="구조화할 최대 건수")
        p.add_argument("--batch", action="store_true", help="Batch API 사용 (50%% 할인, 결과까지 수 분~1시간)")
        p.add_argument("--wait", type=int, default=60, help="배치 완료를 기다릴 최대 시간(분). 넘으면 다음 실행 때 이어받음")
    p = sub.add_parser("review-export")
    p.add_argument("--status", choices=STATUSES, help="이 검토상태인 것만 내보내기")
    p = sub.add_parser("review-import")
    p.add_argument("path", nargs="?", type=Path, help="검토 엑셀 경로 (생략 시 output의 최신 review_*.xlsx)")
    p = sub.add_parser("map")
    p.add_argument("--accounts", required=True, type=Path,
                   help="계정리스트 파일 (분석결과_*.xlsx의 05_계정명 리스트_통계 시트, 또는 계정명 열이 있는 xlsx/csv)")
    p.add_argument("--company", help="회사명 (생략 시 파일에서 읽음)")
    p.add_argument("--no-llm", action="store_true", help="사전 매칭만 수행 (외부 전송 없음)")
    p = sub.add_parser("report")
    p.add_argument("--company", help="회사명 (매칭된 회사가 하나뿐이면 생략 가능)")
    sub.add_parser("status")
    args = parser.parse_args()

    cfg = load_config()
    load_env()
    log = setup_logging(cfg)
    store = Store(cfg["paths"]["db"])
    try:
        if args.command == "extract":
            run_extract(cfg, store, log, force=args.force, name_filter=args.file)
        elif args.command == "segment":
            run_segment(cfg, store, log, force=args.force, name_filter=args.file, force_llm=args.llm)
        elif args.command == "structure":
            run_structure(cfg, store, log, force=args.force, name_filter=args.file, limit=args.limit,
                          batch=args.batch, wait_minutes=args.wait)
        elif args.command == "run-all":
            run_all(cfg, store, log, yes=args.yes, batch=args.batch, limit=args.limit, wait=args.wait)
        elif args.command == "review-export":
            export_review(cfg, store, log, status=args.status)
        elif args.command == "review-import":
            import_review(cfg, store, log, path=args.path.resolve() if args.path else None)
        elif args.command == "map":
            run_map(cfg, store, log, args.accounts.resolve(), company=args.company, use_llm=not args.no_llm)
        elif args.command == "report":
            run_report(cfg, store, log, company=args.company)
        elif args.command == "status":
            show_status(cfg, store, log)
    finally:
        store.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
