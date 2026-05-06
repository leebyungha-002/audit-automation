@echo off
chcp 65001 > nul
cd /d "%~dp0"

if "%~1"=="" (
    echo [ERROR] 회사 폴더명을 인수로 입력하세요.
    echo 사용법: run.bat ^<회사명^> [모드]
    echo 모드: extract ^| inject ^| all  (생략 시 메뉴 선택)
    pause
    exit /b 1
)

set COMPANY=%~1
set MODE=%~2

echo ================================================
echo  Audit Automation : %COMPANY%
echo ================================================

:: ── 환경 활성화 ────────────────────────────────────
if exist "%~dp0activate.bat" (
    echo [ENV] 가상환경 활성화 중...
    call "%~dp0activate.bat"
    goto :select_mode
)

where conda >nul 2>&1
if %ERRORLEVEL%==0 (
    echo [ENV] conda 환경 활성화 중...
    call conda activate audit-automation 2>nul || echo [INFO] conda env 없음 - 기본 환경으로 계속합니다.
)

:select_mode
:: ── 모드 결정 ──────────────────────────────────────
if /i "%MODE%"=="extract" goto :do_extract
if /i "%MODE%"=="inject"  goto :do_inject
if /i "%MODE%"=="all"     goto :do_all

:: 모드 미지정 시 대화형 메뉴
echo.
echo  실행 모드를 선택하세요:
echo  [1] 데이터 추출만    (웹 앱 자동화)
echo  [2] 감사조서 주입만  (data_injector.py)
echo  [3] 순차 실행        (추출 완료 후 자동 주입)
echo.
set /p CHOICE="선택 (1/2/3): "

if "%CHOICE%"=="1" goto :do_extract
if "%CHOICE%"=="2" goto :do_inject
if "%CHOICE%"=="3" goto :do_all

echo [ERROR] 잘못된 입력입니다. 1, 2, 3 중 하나를 입력하세요.
pause
exit /b 1

:: ── 데이터 추출 ────────────────────────────────────
:do_extract
where node >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Node.js가 설치되지 않았거나 PATH에 없습니다.
    pause
    exit /b 1
)
echo.
echo [EXTRACT] 데이터 추출 시작: %COMPANY%
node run.js %COMPANY%
set EXTRACT_CODE=%ERRORLEVEL%
echo.
if %EXTRACT_CODE% NEQ 0 (
    echo [ERROR] 데이터 추출 실패. (Exit code: %EXTRACT_CODE%)
    pause
    exit /b %EXTRACT_CODE%
)
echo [DONE] 데이터 추출 완료: %COMPANY%
if /i "%MODE%"=="extract" goto :end
goto :end

:: ── 감사조서 주입 ───────────────────────────────────
:do_inject
where python >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python이 설치되지 않았거나 PATH에 없습니다.
    pause
    exit /b 1
)
echo.
echo [INJECT] 감사조서 주입 시작: %COMPANY%
python report\data_injector.py %COMPANY%
set INJECT_CODE=%ERRORLEVEL%
echo.
if %INJECT_CODE% NEQ 0 (
    echo [ERROR] 감사조서 주입 실패. (Exit code: %INJECT_CODE%)
    pause
    exit /b %INJECT_CODE%
)
echo [DONE] 감사조서 주입 완료: %COMPANY%
goto :end

:: ── 순차 실행 (추출 → 주입) ────────────────────────
:do_all
where node >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Node.js가 설치되지 않았거나 PATH에 없습니다.
    pause
    exit /b 1
)
where python >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python이 설치되지 않았거나 PATH에 없습니다.
    pause
    exit /b 1
)
echo.
echo [1/2] 데이터 추출 시작: %COMPANY%
node run.js %COMPANY%
set EXTRACT_CODE=%ERRORLEVEL%
echo.
if %EXTRACT_CODE% NEQ 0 (
    echo [ERROR] 데이터 추출 실패 — 주입을 건너뜁니다. (Exit code: %EXTRACT_CODE%)
    pause
    exit /b %EXTRACT_CODE%
)
echo [1/2] 데이터 추출 완료.
echo.
echo [2/2] 감사조서 주입 시작: %COMPANY%
python report\data_injector.py %COMPANY%
set INJECT_CODE=%ERRORLEVEL%
echo.
if %INJECT_CODE% NEQ 0 (
    echo [ERROR] 감사조서 주입 실패. (Exit code: %INJECT_CODE%)
    pause
    exit /b %INJECT_CODE%
)
echo [2/2] 감사조서 주입 완료.
echo.
echo [DONE] 전체 완료: %COMPANY%

:end
pause
