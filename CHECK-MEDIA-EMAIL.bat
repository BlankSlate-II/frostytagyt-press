@echo off
cd /d "%~dp0"
start "" "C:\Users\luked\Desktop\Outlook Classic.bat" 2>nul
timeout /t 5 /nobreak >nul
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0check-media-replies.ps1"
pause