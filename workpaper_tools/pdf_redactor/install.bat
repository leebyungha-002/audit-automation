@echo off
chcp 65001 >nul
set PY=python
if exist "%USERPROFILE%\anaconda3\python.exe" set PY="%USERPROFILE%\anaconda3\python.exe"
echo [설치] %PY% 로 패키지를 설치합니다
%PY% -m pip install -r "%~dp0requirements.txt"
echo.
echo [확인] 설치 상태 점검
%PY% -c "import sys, pymupdf; sys.modules['fitz'] = pymupdf; import pdf2docx, docx, pandas, openpyxl; print('패키지 OK')"
%PY% -c "from winrt.windows.media.ocr import OcrEngine; print('OCR 사용 가능:', ', '.join(l.language_tag for l in OcrEngine.available_recognizer_languages) or '없음 - 한국어 OCR 언어팩 설치 필요')"
pause
