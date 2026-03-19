@echo off
title SpacePad Build Script
echo ============================================
echo   SpacePad Configurator  —  Build Script
echo ============================================
echo.

python --version >nul 2>&1
if errorlevel 1 (echo ERROR: Python not found. & pause & exit /b 1)

echo [1/4] Installing dependencies...
pip install PySide6 pyserial pyinstaller pystray pillow pywin32 --quiet --upgrade
if errorlevel 1 (echo ERROR: pip install failed. & pause & exit /b 1)

echo [2/4] Cleaning previous build...
if exist build rmdir /s /q build
if exist dist  rmdir /s /q dist

echo [3/4] Building SpacePad Configurator.exe...
pyinstaller --noconfirm --onefile --windowed ^
  --name "SpacePad Configurator" ^
  --collect-all PySide6 ^
  --hidden-import serial ^
  --hidden-import serial.tools ^
  --hidden-import serial.tools.list_ports ^
  spacepad_gui.py
if errorlevel 1 (echo ERROR: GUI build failed. & pause & exit /b 1)

echo [4/4] Building SpacePad Tray.exe...
pyinstaller --noconfirm --onefile --windowed ^
  --name "SpacePad Tray" ^
  --hidden-import serial ^
  --hidden-import serial.tools ^
  --hidden-import serial.tools.list_ports ^
  --hidden-import win32gui ^
  --hidden-import pystray ^
  spacepad_tray.py
if errorlevel 1 (echo ERROR: Tray build failed. & pause & exit /b 1)

echo.
echo ============================================
echo   BUILD COMPLETE
echo   SpacePad Configurator.exe  -- main GUI
echo   SpacePad Tray.exe          -- auto-profile switcher
echo   Both are in:  dist\
echo ============================================
echo.
set /p OPEN="Open output folder? (y/n): "
if /i "%OPEN%"=="y" explorer dist
pause
