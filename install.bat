@echo off
title JP-VI Translator Installer
cd /d "%~dp0"

echo ==========================================================
echo    KHOI DONG TRINH CAI DAT JP-VI TRANSLATOR
echo ==========================================================
echo.

:: Check for Administrator privileges
net session >nul 2>&1
if %errorLevel% == 0 (
    echo Quyen Admin da duoc xac nhan. Dang khoi chay PowerShell...
    powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install.ps1"
) else (
    echo Yeu cau quyen Administrator...
    echo Dang yeu cau cap quyen Admin tu he thong (UAC)...
    powershell -Command "Start-Process cmd -ArgumentList '/c \"\"%~0\"\"' -Verb RunAs"
    exit /b
)
