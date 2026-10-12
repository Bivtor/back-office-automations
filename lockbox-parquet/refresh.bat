@echo off
setlocal enabledelayedexpansion
title Lockbox Parquet Refresh
cd /d "%~dp0"

echo ==========================================================
echo   LOCKBOX PARQUET REFRESH
echo ==========================================================
echo.
echo   Converts the Manager Report Detail CSVs into Parquet so
echo   the Power BI reports load faster and use far less memory.
echo.
echo   Run this after new CSVs are dropped in either folder,
echo   BEFORE refreshing the reports in Power BI Desktop.
echo.
echo   Folders handled:
echo     Lockbox 12month raw data   (monthly report^)
echo     Current Raw Data           (weekly report^)
echo     All Data                   (full history^)
echo     UMB daily control sheets   (Control Sheet data^)
echo.
echo   Then recomputes this month's row in KPI Benchmarks.xlsx
echo   and prints the old and new values.
echo.
echo ----------------------------------------------------------
echo.

if not exist "%~dp0lockbox_to_parquet.py" (
    echo   ERROR: lockbox_to_parquet.py is not in this folder.
    echo.
    echo   Both files must sit side by side:
    echo       refresh.bat
    echo       lockbox_to_parquet.py
    echo.
    pause
    exit /b 1
)

REM ---- Locate Python -------------------------------------------------------
REM Actually execute rather than using WHERE: on Windows, "where python" also
REM matches the Microsoft Store app-execution stub, which exists even when
REM Python isn't installed and just opens the Store when run.
set "PY="
py -3 --version >nul 2>&1 && set "PY=py -3"
if not defined PY (
    python --version >nul 2>&1 && set "PY=python"
)

if not defined PY (
    echo   Python was not found on this machine.
    echo.
    echo   To install it:
    echo     1. Open the Microsoft Store
    echo     2. Search for:  Python 3
    echo     3. Install the newest Python 3.x
    echo        ^(publisher should be "Python Software Foundation"^)
    echo     4. Close this window, then run refresh.bat again
    echo.
    echo   No admin rights needed for the Store version.
    echo.
    pause
    exit /b 1
)

echo   Python found:
for /f "delims=" %%v in ('%PY% --version 2^>^&1') do echo     %%v
echo.

REM ---- Dependencies -------------------------------------------------------
echo   Checking required packages ^(pandas, pyarrow, openpyxl^)...
%PY% -c "import pandas, pyarrow, openpyxl" >nul 2>&1
if errorlevel 1 (
    echo   Missing - installing now. This takes a minute the first time only.
    echo.
    %PY% -m pip install --user --upgrade pip
    %PY% -m pip install --user pandas pyarrow openpyxl
    echo.
    %PY% -c "import pandas, pyarrow, openpyxl" >nul 2>&1
    if errorlevel 1 (
        echo   ------------------------------------------------------
        echo   PACKAGE INSTALL FAILED
        echo   ------------------------------------------------------
        echo   Usually a network or proxy block. Try manually:
        echo       %PY% -m pip install --user pandas pyarrow openpyxl
        echo.
        echo   If that is blocked too, IT will need to allow
        echo   pypi.org and files.pythonhosted.org.
        echo.
        pause
        exit /b 1
    )
    echo   Packages installed.
) else (
    echo   Already installed.
)
echo.
echo ----------------------------------------------------------
echo.

REM ---- Run ----------------------------------------------------------------
%PY% "%~dp0lockbox_to_parquet.py" %*
set "RC=%ERRORLEVEL%"

echo.
echo ==========================================================
if "%RC%"=="0" (
    echo   DONE - no errors.
    echo   You can now refresh the reports in Power BI Desktop.
) else (
    echo   FINISHED WITH ERRORS ^(exit code %RC%^)
    echo   Scroll up - the failing file is named in the output.
)
echo ==========================================================
echo.
pause
exit /b %RC%
