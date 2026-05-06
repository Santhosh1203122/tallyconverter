@echo off
REM ============================================================
REM Build TallySalesReport.exe on Windows
REM ============================================================
REM Requirements:
REM   - Python 3.9+ installed and on PATH (https://python.org)
REM   - Internet access for the first run (to install packages)
REM
REM Usage:
REM   1. Open Command Prompt in this folder
REM   2. Run:   build_windows.bat
REM   3. The signed-free executable will appear in dist\TallySalesReport.exe
REM ============================================================

setlocal

echo.
echo === [1/4] Checking Python ===
python --version
if errorlevel 1 (
    echo Python is not installed or not on PATH. Install from https://python.org and tick "Add to PATH".
    exit /b 1
)

echo.
echo === [2/4] Creating virtual environment ===
if not exist .venv (
    python -m venv .venv
)
call .venv\Scripts\activate.bat

echo.
echo === [3/4] Installing dependencies ===
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo.
echo === [4/4] Building EXE ===
REM --onefile : single .exe
REM --windowed: no console window pops up
REM --name    : output name
REM --add-data converter.py;.   : bundle converter alongside app
pyinstaller ^
    --noconfirm ^
    --clean ^
    --onefile ^
    --windowed ^
    --name TallySalesReport ^
    --add-data "converter.py;." ^
    app.py

if exist dist\TallySalesReport.exe (
    echo.
    echo Build successful! See: dist\TallySalesReport.exe
) else (
    echo.
    echo Build failed. Check the output above for errors.
    exit /b 1
)

endlocal
