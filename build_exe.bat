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

for /f "delims=" %%i in ('"%PYTHON_EXE%" -c "import sys; print(sys.base_prefix)"') do set "PY_BASE=%%i"

echo [1/2] Cleaning old build artifacts...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist InvoiceAssistant.spec del /q InvoiceAssistant.spec

echo [2/2] Building EXE with PyInstaller...
"%PYTHON_EXE%" -m PyInstaller --noconfirm --clean --windowed --name InvoiceAssistant ^
  --runtime-hook pyi_rth_tkfix.py ^
  --hidden-import=tkinter ^
  --hidden-import=_tkinter ^
  --add-data "%PY_BASE%\Lib\tkinter;tkinter" ^
  --add-data "%PY_BASE%\tcl\tcl8.6;_tcl_data" ^
  --add-data "%PY_BASE%\tcl\tk8.6;_tk_data" ^
  --add-binary "%PY_BASE%\DLLs\_tkinter.pyd;." ^
  --add-binary "%PY_BASE%\DLLs\tcl86t.dll;." ^
  --add-binary "%PY_BASE%\DLLs\tk86t.dll;." ^
  --collect-all pypdfium2 ^
  --collect-all cv2 ^
  --collect-all rapidocr_onnxruntime ^
  --collect-all onnxruntime ^
  app.py

if %ERRORLEVEL% neq 0 (
  echo Build failed.
  exit /b %ERRORLEVEL%
)

echo Build success.
echo EXE path: %WORKDIR%dist\InvoiceAssistant\InvoiceAssistant.exe
endlocal
