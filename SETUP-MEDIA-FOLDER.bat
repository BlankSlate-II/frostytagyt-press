@echo off
cd /d "%~dp0"
start "" "%~dp0media-inbox-setup.html"
start "" "https://outlook.live.com/mail/0/options/mail/rules"
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0open-outlook-rules.ps1"