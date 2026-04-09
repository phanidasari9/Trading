@echo off
setlocal
cd /d "%~dp0"
if not exist ".venv\Scripts\python.exe" (
  echo Creating .venv ...
  python -m venv .venv
)
".venv\Scripts\python.exe" -m pip install --upgrade pip
".venv\Scripts\python.exe" -m pip install -e .
echo.
echo Build finished. Run start_dashboard.bat or: .venv\Scripts\python.exe -m streamlit run dashboard.py
endlocal
