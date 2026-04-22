@echo off
chcp 65001 > nul
cd /d "%~dp0"

if "%~1"=="" (
    echo [ERROR] Please provide the company folder name as an argument.
    echo Usage: run.bat Braintree
    pause
    exit /b 1
)

set COMPANY=%~1
echo ================================================
echo  Audit Automation : %COMPANY%
echo ================================================

if exist "%~dp0activate.bat" (
    echo [ENV] Activating virtual environment...
    call "%~dp0activate.bat"
    goto :run
)

where conda >nul 2>&1
if %ERRORLEVEL%==0 (
    echo [ENV] Activating conda environment...
    call conda activate audit-automation 2>nul || echo [INFO] No conda env - continuing with default.
)

:run
where node >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Node.js is not installed or not in PATH.
    pause
    exit /b 1
)

echo [RUN] node run.js %COMPANY%
echo.
node run.js %COMPANY%

echo.
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Execution failed. Check logs above. (Exit code: %ERRORLEVEL%)
) else (
    echo [DONE] %COMPANY% completed successfully.
)

pause
