"""감리 지적사항 ↔ 계정리스트 대사 체크리스트 생성기 CLI.

사용 예:
    python main.py extract                 # 원문(PDF/HWP) 텍스트 추출
    python main.py segment                 # 지적사항 단위 분할
    python main.py structure --limit 5     # Claude API로 구조화 (처음 5건만)
    python main.py review-export           # 검토용 엑셀 내보내기 (output/review_날짜.xlsx)
    python main.py review-import           # 수정한 검토 엑셀을 DB에 재반영 (경로 생략 시 최신 파일)
    python main.py map --accounts <파일>   # 계정리스트를 표준 분류로 정규화하고 지적사항과 매칭
    python main.py status                  # 현황 요약
"""
import argparse
import sys
from pathlib import Path

from afc.config import load_config, load_env, setup_logging
from afc.db import Store
from afc.extract import run_extract
from afc.mapping import run_map
from afc.review import STATUSES, export_review, import_review
from afc.segment import run_segment
from afc.structure import run_structure


def show_status(store: Store) -> None:
    print("\n[원문 파일]")
    for r in store.query(
        "SELECT f.file_name, f.issuer, f.n_pages, f.extract_mode, "
        "(SELECT COUNT(*) FROM segments s WHERE s.file_hash=f.file_hash) AS n_seg, "
        "(SELECT COUNT(*) FROM findings d JOIN segments s ON s.segment_id=d.segment_id "
        " WHERE s.file_hash=f.file_hash) AS n_find FROM source_files f ORDER BY f.file_name"
    ):
        print(f"  {r['file_name']} | {r['issuer']} | {r['n_pages']}쪽 | {r['extract_mode']} | "
              f"분할 {r['n_seg']}건 | 구조화 {r['n_find']}건")
    print("\n[검토 상태]")
    for r in store.query("SELECT review_status, COUNT(*) AS n FROM findings GROUP BY review_status"):
        print(f"  {r['review_status']}: {r['n']}건")
    print("\n[발췌 검증]")
    for r in store.query("SELECT excerpt_status, COUNT(*) AS n FROM findings GROUP BY excerpt_status"):
        print(f"  {r['excerpt_status']}: {r['n']}건")
    print("\n[경고/오류 로그]")
    for r in store.query("SELECT level, COUNT(*) AS n FROM process_log WHERE level<>'INFO' GROUP BY level"):
        print(f"  {r['level']}: {r['n']}건")


def main() -> int:
    parser = argparse.ArgumentParser(description="감리 지적사항 체크리스트 생성기")
    sub = parser.add_subparsers(dest="command", required=True)
    for name in ("extract", "segment", "structure"):
        p = sub.add_parser(name)
        p.add_argument("--force", action="store_true", help="이미 처리한 것도 다시 처리")
        p.add_argument("--file", help="파일명에 이 문자열이 포함된 것만 처리")
        if name == "segment":
            p.add_argument("--llm", action="store_true", help="규칙 대신 LLM 보조 분할 강제")
        if name == "structure":
            p.add_argument("--limit", type=int, help="처리할 최대 건수")
    p = sub.add_parser("review-export")
    p.add_argument("--status", choices=STATUSES, help="이 검토상태인 것만 내보내기")
    p = sub.add_parser("review-import")
    p.add_argument("path", nargs="?", type=Path, help="검토 엑셀 경로 (생략 시 output의 최신 review_*.xlsx)")
    p = sub.add_parser("map")
    p.add_argument("--accounts", required=True, type=Path,
                   help="계정리스트 파일 (분석결과_*.xlsx의 05_계정명 리스트_통계 시트, 또는 계정명 열이 있는 xlsx/csv)")
    p.add_argument("--company", help="회사명 (생략 시 파일에서 읽음)")
    p.add_argument("--no-llm", action="store_true", help="사전 매칭만 수행 (외부 전송 없음)")
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
            run_structure(cfg, store, log, force=args.force, name_filter=args.file, limit=args.limit)
        elif args.command == "review-export":
            export_review(cfg, store, log, status=args.status)
        elif args.command == "review-import":
            import_review(cfg, store, log, path=args.path.resolve() if args.path else None)
        elif args.command == "map":
            run_map(cfg, store, log, args.accounts.resolve(), company=args.company, use_llm=not args.no_llm)
        elif args.command == "status":
            show_status(store)
    finally:
        store.close()
    return 0


if __name__ == "__main__":
    sys.exit(main())
