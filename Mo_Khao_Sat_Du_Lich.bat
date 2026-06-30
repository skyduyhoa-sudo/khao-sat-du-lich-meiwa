@echo off
setlocal
cd /d "%~dp0"

set "PYTHON_EXE=D:\hop dong thuy tro ban word\aI BUIDER\bulid app by codex\qc-spc-dashboard\.venv\Scripts\python.exe"

if not exist "%PYTHON_EXE%" (
  echo Khong tim thay Python trong environment da cau hinh.
  echo Hay kiem tra lai duong dan: %PYTHON_EXE%
  pause
  exit /b 1
)

start "" "%PYTHON_EXE%" -m streamlit run app.py
exit /b 0
