@echo off
chcp 65001 >nul
"C:\Users\USer\anaconda3\python.exe" "%~dp0transcribe.py" %*
pause
