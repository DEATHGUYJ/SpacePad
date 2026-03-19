@echo off
title SpacePad Build Script
echo ============================================
echo   SpacePad Configurator  —  Build Script
echo ============================================
echo.

python --version >nul 2>&1
if errorlevel 1 (echo ERROR: Python not found. & pause & exit /b 1)

echo [1/3] Installing dependencies...
pip install PySide6 pyserial psutil pyinstaller --quiet --upgrade
if errorlevel 1 (echo ERROR: pip install failed. & pause & exit /b 1)

echo [2/3] Cleaning previous build...
if exist build rmdir /s /q build
if exist dist  rmdir /s /q dist

echo [3/3] Building SpacePad Configurator.exe...
pyinstaller --noconfirm --onefile --windowed ^
  --name "SpacePad Configurator" ^
  --collect-all PySide6 ^
  --hidden-import serial ^
  --hidden-import serial.tools ^
  --hidden-import serial.tools.list_ports ^
  --hidden-import psutil ^
  spacepad_gui.py
if errorlevel 1 (echo ERROR: Build failed. & pause & exit /b 1)

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
