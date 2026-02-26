@echo off
setlocal
set SCRIPT_DIR=%~dp0
if not exist "%SCRIPT_DIR%\.venv\Scripts\python.exe" (
  echo Virtual environment not found. Run install.ps1 first.
  exit /b 1
)
"%SCRIPT_DIR%\.venv\Scripts\python.exe" "%SCRIPT_DIR%\emis_letter_summary_ui.py"
endlocal
