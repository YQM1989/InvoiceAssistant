@echo off
setlocal
set "PYTHON_EXE=%~dp0.venv\Scripts\python.exe"
if exist "%PYTHON_EXE%" (
  "%PYTHON_EXE%" "%~dp0app.py"
) else (
  echo [WARN] Virtual environment not found. Trying system Python...
  py "%~dp0app.py"
)
endlocal
