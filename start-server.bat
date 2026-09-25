@echo off
title Book Barcode Scanner Local Server
echo =================================================================
echo   Starting Book Barcode Scanner and Excel Export System...
echo =================================================================
echo.
echo Launching local server at http://localhost:8080 ...
echo (This enables camera permissions in a secure localhost context)
echo.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0start-server.ps1"
pause
