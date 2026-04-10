@echo off
cd /d "%~dp0"
echo 감사 자동화 툴을 시작합니다...
call npm run dev %1
pause
