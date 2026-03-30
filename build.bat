@echo off
title SpacePad GUI Builder

echo ============================================
echo   SpacePad Configurator  —  Build Script
echo ============================================
echo.

REM ── Find Python ──────────────────────────────────────────
set "PYTHON="

REM Try 'py' launcher first
py --version >nul 2>&1
if not errorlevel 1 (
    set "PYTHON=py"
    goto :found
)

REM Try 'python' on PATH
python --version >nul 2>&1
if not errorlevel 1 (
    set "PYTHON=python"
    goto :found
)

REM Search common install locations
for %%D in (
    "%LOCALAPPDATA%\python\pythoncore-3.14-64"
    "%LOCALAPPDATA%\python\pythoncore-3.13-64"
    "%LOCALAPPDATA%\python\pythoncore-3.12-64"
    "%LOCALAPPDATA%\Programs\Python\Python314"
    "%LOCALAPPDATA%\Programs\Python\Python313"
    "%LOCALAPPDATA%\Programs\Python\Python312"
    "%ProgramFiles%\Python314"
    "%ProgramFiles%\Python313"
    "%ProgramFiles%\Python312"
) do (
    if exist "%%~D\python.exe" (
        set "PYTHON=%%~D\python.exe"
        goto :found
    )
)

echo ERROR: Python not found.
echo.
echo   Install Python from https://python.org
echo   Make sure to check "Add Python to PATH" during install.
echo.
echo   If Python is already installed, edit this script and set
echo   PYTHON= to your python.exe path, for example:
echo   set "PYTHON=C:\Users\you\AppData\Local\python\...\python.exe"
pause
exit /b 1

:found
echo Found Python: %PYTHON%
%PYTHON% --version
echo.

REM ── Install / upgrade dependencies ──
echo [1/3] Installing dependencies...
%PYTHON% -m pip install PySide6 pyserial psutil pyinstaller --quiet --upgrade
if errorlevel 1 (
    echo.
    echo ERROR: pip install failed.
    echo   Try running this script as Administrator, or run manually:
    echo   %PYTHON% -m pip install PySide6 pyserial psutil pyinstaller
    pause
    exit /b 1
)

REM ── Clean previous build ──
echo [2/3] Cleaning previous build...
if exist build rmdir /s /q build
if exist dist  rmdir /s /q dist

REM ── Build the exe ──
echo [3/3] Building SpacePad Configurator.exe...
%PYTHON% -m PyInstaller --noconfirm --onefile --windowed ^
  --name "SpacePad Configurator" ^
  --collect-all PySide6 ^
  --hidden-import serial ^
  --hidden-import serial.tools ^
  --hidden-import serial.tools.list_ports ^
  --hidden-import psutil ^
  spacepad_gui.py
if errorlevel 1 (
    echo.
    echo ERROR: Build failed. See output above.
    pause
    exit /b 1
)

echo.
echo ============================================
echo   BUILD COMPLETE
echo   SpacePad Configurator.exe is in:  dist\
echo.
echo   Features:
echo   - Full GUI configurator
echo   - Minimizes to system tray
echo   - Auto layer switching by app
echo ============================================
echo.
set /p OPEN="Open output folder? (y/n): "
if /i "%OPEN%"=="y" explorer dist
pause
