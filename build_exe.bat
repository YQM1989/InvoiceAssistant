@echo off
setlocal
set "WORKDIR=%~dp0"
cd /d "%WORKDIR%"

set "PYTHON_EXE=%WORKDIR%.venv\Scripts\python.exe"
if not exist "%PYTHON_EXE%" (
  echo [ERROR] Virtual environment not found: %PYTHON_EXE%
  echo Please create a virtual environment first: python -m venv .venv
  echo Then install dependencies: .venv\Scripts\python.exe -m pip install -r requirements.txt
  exit /b 1
)

echo [1/2] Cleaning old build artifacts...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist

echo [2/2] Building EXE with PyInstaller...
"%PYTHON_EXE%" -m PyInstaller --noconfirm --clean InvoiceAssistant.spec

if %ERRORLEVEL% neq 0 (
  echo Build failed.
  exit /b %ERRORLEVEL%
)

echo Build success.
echo EXE path: %WORKDIR%dist\InvoiceAssistant\InvoiceAssistant.exe
endlocal
