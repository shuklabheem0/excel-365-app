@echo off
title Excel 365 Desktop - Windows Installer Builder
echo ========================================================
echo   Excel 365 Offline Desktop Installer (.exe) Builder
echo ========================================================
echo.

where node >nul 2>nul
if %errorlevel% neq 0 (
    echo [ERROR] Node.js aapke PC me install nahi hai!
    echo Pehle https://nodejs.org/ se LTS version install karein.
    pause
    exit /b
)

echo [1/2] Zaroori packages install ho rahe hain...
call npm install
if %errorlevel% neq 0 (
    echo [ERROR] npm install fail ho gaya. Internet connection check karein.
    pause
    exit /b
)

echo.
echo [2/2] Windows NSIS Installer (.exe) compile ho raha hai...
call npx vite build && call npx electron-builder --win nsis portable --x64
if %errorlevel% neq 0 (
    echo [ERROR] Installer build fail ho gaya.
    pause
    exit /b
)

echo.
echo ========================================================
echo   SUCCESS! Installer taiyar hai.
echo   Folder: dist_electron\Excel 365 Desktop Setup 1.0.0.exe
echo ========================================================
start "" "dist_electron"
pause
