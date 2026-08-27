@echo off
setlocal

cd /d "%~dp0"
set LOGFILE=%~dp0sync_pull.log

echo [%date% %time%] git pull 시작 >> "%LOGFILE%"
git pull --ff-only >> "%LOGFILE%" 2>&1

if %ERRORLEVEL% NEQ 0 (
    echo [%date% %time%] git pull 실패 - 로컬 변경/충돌 또는 네트워크 오류일 수 있음, 수동 확인 필요 >> "%LOGFILE%"
) else (
    echo [%date% %time%] git pull 완료 >> "%LOGFILE%"
)
echo. >> "%LOGFILE%"

endlocal
