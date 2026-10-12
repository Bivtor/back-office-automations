@echo off
REM ============================================================
REM  Outstanding Rollover Report - launcher
REM  Double-click this. No admin, no install.
REM  Starts server.ps1 (local web server) then opens the GUI
REM  after a short delay to let the listener bind first.
REM
REM  Files expected next to this one:
REM    server.ps1   index.html   run.ps1   config.json
REM ============================================================

cd /d "%~dp0"

if not exist "%~dp0server.ps1" (
  echo ERROR: server.ps1 not found next to START.bat
  pause
  exit /b 1
)
if not exist "%~dp0run.ps1" (
  echo ERROR: run.ps1 not found next to START.bat
  pause
  exit /b 1
)
if not exist "%~dp0config.json" (
  echo WARNING: config.json not found next to START.bat - the run will fail.
  echo.
)

REM Give the listener a moment to bind before the browser hits it.
start "" /B powershell -NoProfile -ExecutionPolicy Bypass -Command ^
  "Start-Sleep -Milliseconds 900; Start-Process 'http://localhost:8080/'"

powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0server.ps1"

echo.
echo Server has stopped. You can close this window.
pause
