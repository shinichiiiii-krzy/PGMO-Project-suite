@echo off
setlocal

echo ============================================================
echo  PGMO Project Suite - Build Script
echo ============================================================
echo.

REM --- Check Python is available ---
where python >nul 2>nul
if errorlevel 1 (
    echo [ERROR] Python was not found on this machine.
    echo Install Python 3.10+ from https://www.python.org/downloads/
    echo and make sure to check "Add python.exe to PATH" during setup.
    pause
    exit /b 1
)

echo [1/4] Creating a clean virtual environment (.buildenv)...
if exist .buildenv (
    echo    .buildenv already exists, reusing it.
) else (
    python -m venv .buildenv
)
call .buildenv\Scripts\activate.bat

echo.
echo [2/4] Installing dependencies (pandas, openpyxl, xlrd, pypdf, pyinstaller)...
python -m pip install --upgrade pip >nul
pip install -r requirements.txt
if errorlevel 1 (
    echo [ERROR] pip install failed. Check your internet connection and try again.
    pause
    exit /b 1
)

echo.
echo [3/4] Cleaning previous build output...
if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist "PGMO Project Suite.spec" del /q "PGMO Project Suite.spec"

echo.
echo [4/4] Building the standalone .exe with PyInstaller...
pyinstaller --noconfirm --onefile --windowed ^
    --name "PGMO Project Suite" ^
    --icon "app_icon.ico" ^
    --add-data "app_icon.ico;." ^
    --hidden-import "xlrd" ^
    --collect-submodules "xlrd" ^
    PGMO_PROJECT_GUI_MERGE_CSV.py

if errorlevel 1 (
    echo.
    echo [ERROR] Build failed. Scroll up to see the PyInstaller error.
    pause
    exit /b 1
)

echo.
echo ============================================================
echo  BUILD COMPLETE
echo  Your app is at: dist\PGMO Project Suite.exe
echo  You can copy that single .exe anywhere and double-click it -
echo  no Python installation needed on the target machine.
echo ============================================================
echo.
pause
