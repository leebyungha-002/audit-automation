"""런처용 실행 스크립트: Streamlit 화면을 띄우고 브라우저를 자동으로 연다.

- 같은 프로세스에서 Streamlit을 실행하므로 런처의 [중지]로 서버까지 함께 종료된다.
- headless 모드라 최초 실행 시 이메일 입력 대기(stdin)로 멈추지 않는다.
"""
import sys
import threading
import webbrowser
from pathlib import Path

PORT = 8510
APP = Path(__file__).with_name("app.py")


def main() -> None:
    url = f"http://localhost:{PORT}"
    print(f"FSS 감사 참고자료 검색 화면을 엽니다: {url}")
    print("종료하려면 런처의 [중지]를 누르세요.")
    threading.Timer(3, webbrowser.open, args=(url,)).start()
    from streamlit.web import cli
    sys.argv = ["streamlit", "run", str(APP), "--server.headless", "true",
                "--server.port", str(PORT), "--browser.gatherUsageStats", "false"]
    sys.exit(cli.main())


if __name__ == "__main__":
    main()
