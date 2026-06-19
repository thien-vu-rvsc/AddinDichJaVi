@echo off
title Cai dat JP-VI Translator Add-in
echo Dang tien hanh cai dat Add-in dich thuat JP-VI...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install-client.ps1"
echo.
echo Nhan phim bat ky de dong...
pause > nul
