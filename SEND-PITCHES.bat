@echo off
title FrostyTagYT - Send All Pitches
echo.
echo  This opens all 4 pitches automatically.
echo  You only need to:
echo    1. Click SEND on the Road to VR email
echo    2. Paste (Ctrl+V) and submit on each of the 3 web forms
echo.
powershell -ExecutionPolicy Bypass -File "%~dp0send-all-pitches.ps1"
echo.
pause