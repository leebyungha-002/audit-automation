"""사용법:
  python main.py sync 2018 2026            (전체 자료)
  python main.py sync 2018 2026 주석공시    (특정 자료만)
  python main.py search 1118
  python main.py apis                      (자료 구분 목록)
"""
import os
import sys

from dotenv import load_dotenv

import fss_client
import store


def main(argv: list[str]) -> None:
    con = store.connect(os.getenv("FSS_DB_PATH", "fss.db"))
    if len(argv) >= 3 and argv[0] == "sync":
        load_dotenv()
        key = os.getenv("FSS_API_KEY")
        if not key:
            sys.exit(".env 파일에 FSS_API_KEY를 설정하세요.")
        apis = argv[3:] or None
        unknown = [a for a in apis or [] if a not in fss_client.ENDPOINTS]
        if unknown:
            sys.exit(f"알 수 없는 자료: {', '.join(unknown)} (python main.py apis 로 확인)")
        failed = 0
        for api, year, got, changed, err in store.collect(con, key, int(argv[1]), int(argv[2]), apis):
            if err:
                failed += 1
                print(f"{api} {year}: 실패 - {err} (이 자료는 건너뜀)")
            else:
                print(f"{api} {year}: 수신 {got}건, 신규/변경 {changed}건")
        if failed:
            sys.exit(1)
    elif len(argv) == 2 and argv[0] == "search":
        rows = store.search(con, argv[1])
        for api, reg, subject, url, fetched, body in rows:
            print(f"[{api}] {(reg or '')[:10]} {subject}\n  {store.snippet(body, argv[1])}"
                  f"\n  원문: {url}\n  수집: {fetched}")
        print(f"총 {len(rows)}건")
    elif argv == ["apis"]:
        print("\n".join(fss_client.ENDPOINTS))
    else:
        print(__doc__)


if __name__ == "__main__":
    main(sys.argv[1:])
