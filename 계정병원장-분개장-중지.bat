@echo off
cd /d "%~dp0"
echo 실행 중인 자동화 노드 프로세스(Playwright 등)를 모두 강제 종료합니다...
taskkill /F /IM node.exe /T
echo.
echo 모든 종료 작업이 완료되었습니다.
pause
