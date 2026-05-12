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
echo  [1] 데이터 추출만    (run.js + interest_expense_extractor.js)
echo  [2] 감사조서 주입만  (data_injector.py + interest_expense_analysis.py)
echo  [3] 순차 실행        (추출 완료 후 자동 주입 + 분석)
echo.
set /p CHOICE="선택 (1/2/3): "

if "%CHOICE%"=="1" goto :do_extract
if "%CHOICE%"=="2" goto :do_inject
if "%CHOICE%"=="3" goto :do_all

echo [ERROR] 잘못된 입력입니다. 1, 2, 3 중 하나를 입력하세요.
pause
exit /b 1

:: ── 데이터 추출 ─────────────────────────────────────────────────────────────
:do_extract
where node >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Node.js가 설치되지 않았거나 PATH에 없습니다.
    pause
    exit /b 1
)

echo.
echo [1/3] 메인 데이터 추출 시작: %COMPANY%
node run.js %COMPANY%
set EXTRACT_CODE=%ERRORLEVEL%
echo.
if %EXTRACT_CODE% NEQ 0 (
    echo [ERROR] 메인 데이터 추출 실패. (Exit code: %EXTRACT_CODE%)
    pause
    exit /b %EXTRACT_CODE%
)
echo [1/3] 메인 데이터 추출 완료.

echo.
echo [2/3] 상세검색 직접 추출 시작: %COMPANY%
python detail_search_extractor.py %COMPANY%
set DSE_CODE=%ERRORLEVEL%
echo.
if %DSE_CODE% NEQ 0 (
    echo [WARN] 상세검색 추출 실패 — 계속 진행합니다. (Exit code: %DSE_CODE%)
)
echo [2/3] 상세검색 직접 추출 완료.

echo.
echo [3/3] 이자비용 원장 추출 시작: %COMPANY%
node interest_expense_extractor.js %COMPANY%
set IEE_CODE=%ERRORLEVEL%
echo.
if %IEE_CODE% NEQ 0 (
    echo [WARN] 이자비용 원장 추출 실패 — 계속 진행합니다. (Exit code: %IEE_CODE%)
)
echo [3/3] 이자비용 원장 추출 완료.

echo.
echo [DONE] 데이터 추출 완료: %COMPANY%
goto :end

:: ── 감사조서 주입 + 적정성 분석 ─────────────────────────────────────────────
:do_inject
where python >nul 2>&1
if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Python이 설치되지 않았거나 PATH에 없습니다.
    pause
    exit /b 1
)

echo.
echo [1/2] 감사조서 주입 시작: %COMPANY%
python report\data_injector.py %COMPANY%
set INJECT_CODE=%ERRORLEVEL%
echo.
if %INJECT_CODE% NEQ 0 (
    echo [ERROR] 감사조서 주입 실패. (Exit code: %INJECT_CODE%)
    pause
    exit /b %INJECT_CODE%
)
echo [1/2] 감사조서 주입 완료.

echo.
echo [2/2] 이자비용 적정성 분석 시작: %COMPANY%
python interest_expense_analysis.py %COMPANY%\results\이자비용적정성.xlsx
set ANALYSIS_CODE=%ERRORLEVEL%
echo.
if %ANALYSIS_CODE% NEQ 0 (
    echo [WARN] 이자비용 적정성 분석 실패 — 계속 진행합니다. (Exit code: %ANALYSIS_CODE%)
)
echo [2/2] 이자비용 적정성 분석 완료.

echo.
echo [DONE] 주입 및 분석 완료: %COMPANY%
goto :end

:: ── 순차 실행 (추출 → 상세검색 → 주입 → 이자비용추출 → 분석) ─────────────
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
echo [1/5] 메인 데이터 추출 시작: %COMPANY%
node run.js %COMPANY%
set EXTRACT_CODE=%ERRORLEVEL%
echo.
if %EXTRACT_CODE% NEQ 0 (
    echo [ERROR] 메인 데이터 추출 실패 — 이후 단계를 건너뜁니다. (Exit code: %EXTRACT_CODE%)
    pause
    exit /b %EXTRACT_CODE%
)
echo [1/5] 메인 데이터 추출 완료.

echo.
echo [2/5] 상세검색 직접 추출 시작: %COMPANY%
python detail_search_extractor.py %COMPANY%
set DSE_CODE=%ERRORLEVEL%
echo.
if %DSE_CODE% NEQ 0 (
    echo [WARN] 상세검색 추출 실패 — 계속 진행합니다. (Exit code: %DSE_CODE%)
)
echo [2/5] 상세검색 직접 추출 완료.

echo.
echo [3/5] 감사조서 주입 시작: %COMPANY%
python report\data_injector.py %COMPANY%
set INJECT_CODE=%ERRORLEVEL%
echo.
if %INJECT_CODE% NEQ 0 (
    echo [ERROR] 감사조서 주입 실패. (Exit code: %INJECT_CODE%)
    pause
    exit /b %INJECT_CODE%
)
echo [3/5] 감사조서 주입 완료.

echo.
echo [4/5] 이자비용 원장 추출 시작: %COMPANY%
node interest_expense_extractor.js %COMPANY%
set IEE_CODE=%ERRORLEVEL%
echo.
if %IEE_CODE% NEQ 0 (
    echo [WARN] 이자비용 원장 추출 실패 — 계속 진행합니다. (Exit code: %IEE_CODE%)
)
echo [4/5] 이자비용 원장 추출 완료.

echo.
echo [5/5] 이자비용 적정성 분석 시작: %COMPANY%
python interest_expense_analysis.py %COMPANY%\results\이자비용적정성.xlsx
set ANALYSIS_CODE=%ERRORLEVEL%
echo.
if %ANALYSIS_CODE% NEQ 0 (
    echo [WARN] 이자비용 적정성 분석 실패 — 계속 진행합니다. (Exit code: %ANALYSIS_CODE%)
)
echo [5/5] 이자비용 적정성 분석 완료.

echo.
echo [DONE] 전체 완료: %COMPANY%

:end
pause
