@echo off
REM Daily control sheet run. Double-click this file.
cd /d "%~dp0"

where py >nul 2>nul && (set PY=py) || (set PY=python)

%PY% server.py
if errorlevel 1 (
  echo.
  echo Server stopped with an error. Check the message above.
  pause
)
