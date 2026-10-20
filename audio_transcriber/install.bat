@echo off
chcp 65001 >nul
set PY=python
if exist "%USERPROFILE%\anaconda3\python.exe" set PY="%USERPROFILE%\anaconda3\python.exe"
echo [설치] %PY% 로 패키지를 설치합니다
%PY% -m pip install -r "%~dp0requirements.txt"
pause
