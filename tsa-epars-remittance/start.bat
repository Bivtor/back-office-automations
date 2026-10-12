@echo off
title TPA Remittance Processing - Setup
cd /d "%~dp0"

REM Check if generate_command.ps1 exists
if not exist "generate_command.ps1" (
    echo ERROR: generate_command.ps1 not found in current directory!
    echo Current directory: %CD%
    echo.
    pause
    exit /b 1
)

REM Run the PowerShell script
powershell.exe -ExecutionPolicy Bypass -File "%~dp0generate_command.ps1"

pause
