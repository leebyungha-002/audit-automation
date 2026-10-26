@echo off
chcp 65001 >nul
set PY=python
if exist "%USERPROFILE%\anaconda3\python.exe" set PY="%USERPROFILE%\anaconda3\python.exe"
%PY% "%~dp0redact_pdf.py" %*
pause
