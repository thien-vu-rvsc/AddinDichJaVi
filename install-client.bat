@echo off
chcp 65001 > nul
title Cài đặt JP-VI Translator Add-in
echo Đang tiến hành cài đặt Add-in dịch thuật JP-VI...
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install-client.ps1"
echo.
echo Nhấn phím bất kỳ để đóng...
pause > nul
