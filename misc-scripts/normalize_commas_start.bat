@echo off
REM Launches normalize_dollars.ps1 from the same folder as this .bat
REM Bypasses execution policy for this single run only.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0normalize_commas.ps1"
