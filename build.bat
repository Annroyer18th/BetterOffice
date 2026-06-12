@echo off
cd /d "%~dp0"

echo ============================================
echo   Excel Merger - Build Script
echo ============================================
echo.

python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python not found. Please install Python 3.8+
    echo         https://www.python.org/downloads/
    pause
    exit /b 1
)

echo [1/3] Installing dependencies...
pip install openpyxl xlrd pywin32 pyinstaller -q
if %errorlevel% neq 0 (
    echo [ERROR] Failed to install dependencies
    pause
    exit /b 1
)

echo [2/3] Cleaning old builds...
if exist "dist" rmdir /s /q "dist"
if exist "build" rmdir /s /q "build"
if exist "*.spec" del /q "*.spec"

echo [3/3] Building single exe...
pyinstaller --onefile --windowed --name "ExcelMerger" --hidden-import openpyxl.cell._writer --hidden-import openpyxl.styles --hidden-import xlrd --hidden-import win32com --clean excel_merger.py

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Build failed. See errors above.
    pause
    exit /b 1
)

echo.
echo ============================================
echo   Build SUCCESS!
echo   Output: dist\ExcelMerger.exe
echo ============================================
echo.
echo You can run ExcelMerger.exe directly - no install needed.
echo.

set /p open="Open output folder? (y/n): "
if /i "%open%"=="y" start "" "%~dp0dist"

pause
