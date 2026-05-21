@echo off
cd /d "%~dp0"
echo.
echo  スケジュール調整アプリを起動します...
echo.
start "" http://localhost:3000
node server.js
pause
