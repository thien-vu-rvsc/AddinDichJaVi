@echo off
title JP-VI Translator Backend Server
cd /d "%~dp0"
echo ==========================================================
echo    KHOI DONG SERVER DICH THUAT JP-VI CUC BO
echo ==========================================================
echo.

if exist agent_backend.exe (
    echo Dang chay tu file thuc thi agent_backend.exe...
    agent_backend.exe --prod
) else (
    echo Dang chay tu file script python agent_backend.py...
    python agent_backend.py --prod
)

echo.
echo Server da dung hoac xay ra loi.
pause
