@echo off
REM ============================================================
REM  Daily Queue Processor - launcher
REM  Double-click this. No admin, no install.
REM  Starts server.ps1 (local web server) then opens the GUI
REM  after a short delay to let the listener bind first.
REM ============================================================

cd /d "%~dp0"

REM Wait ~1.5s for the listener to bind before opening the browser.
REM (ping -n 3 gives ~2 second delay without needing timeout.exe)
start "" /B powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Start-Sleep -Milliseconds 150; Start-Process 'http://localhost:8080/'"

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0server.ps1"

echo.
echo Server has stopped. You can close this window.
pause
