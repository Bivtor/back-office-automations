# Weekly Rollover Report - Command Generator
# This script helps users configure and run the weekly rollover processor

Write-Host "=== Weekly Rollover Report - Command Generator ===" -ForegroundColor Green
Write-Host ""

# Open KB article automatically before asking any questions
Write-Host "Opening knowledge base article..." -ForegroundColor Cyan
Start-Process "https://kb.example.com/knowledgebase/weekly-rollover-report-transfers-in/#:~:text=Building%20the%20Reports-,Download%20and%20Save%20Report%20Data,-Navigate%20to%20the"
Write-Host ""

# Get script location
$ScriptPath = Split-Path -Parent $MyInvocation.MyCommand.Path
$RunScriptPath = Join-Path $ScriptPath "run.ps1"

# Verify run.ps1 exists
if (-not (Test-Path $RunScriptPath)) {
    Write-Host "ERROR: run.ps1 not found at: $RunScriptPath" -ForegroundColor Red
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit
}

# Verify config.json exists
$ConfigPath = Join-Path $ScriptPath "config.json"
if (-not (Test-Path $ConfigPath)) {
    Write-Host "ERROR: config.json not found at: $ConfigPath" -ForegroundColor Red
    Write-Host "Please ensure config.json is in the same folder as this script." -ForegroundColor Yellow
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit
}

# Validate config.json is parseable
try {
    $Config = Get-Content $ConfigPath | ConvertFrom-Json
    Write-Host "Config file loaded successfully." -ForegroundColor Green
    Write-Host "  Output base path : $($Config.outputBasePath)" -ForegroundColor Cyan
    Write-Host ""
}
catch {
    Write-Host "ERROR: config.json is not valid JSON: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit
}

# Get current username for Downloads path
$CurrentUser = $env:USERNAME
$AssumedDownloads = "C:\Users\$CurrentUser\Downloads"

# -----------------------------
# STEP 1: Raw Input File
# -----------------------------

# Glob for newest matching file in Downloads
$AssumedInputFile = Get-ChildItem $AssumedDownloads -File |
    Where-Object { $_.Name -match '(?i)Custom\s+Rollover\s+Search\s+Results.*\.xlsx$' } |
    Sort-Object LastWriteTime -Descending |
    Select-Object -First 1

$RawInputValid = $false
while (-not $RawInputValid) {
    Write-Host "Step 1: Raw Input File (Custom Rollover Search Results.xlsx)" -ForegroundColor Yellow

    if ($AssumedInputFile) {
        Write-Host "Assumed file: $($AssumedInputFile.FullName)" -ForegroundColor Cyan
    } else {
        Write-Host "No matching file found automatically in: $AssumedDownloads" -ForegroundColor Cyan
    }

    Write-Host "Press Enter to accept, or paste a different path:"
    $RawInputInput = Read-Host "Raw input file path"

    if ([string]::IsNullOrWhiteSpace($RawInputInput)) {
        if ($AssumedInputFile) {
            $RawInputFile = $AssumedInputFile.FullName
        } else {
            Write-Host ""
            Write-Host "ERROR: No file was auto-detected and no path was entered." -ForegroundColor Red
            Write-Host ""
            continue
        }
    } else {
        $RawInputFile = $RawInputInput.Trim().Trim('"', "'")
    }

    if (Test-Path $RawInputFile) {
        Write-Host "File verified successfully!" -ForegroundColor Green
        $RawInputValid = $true
    } else {
        Write-Host ""
        Write-Host "ERROR: File does not exist: $RawInputFile" -ForegroundColor Red
        Write-Host ""
    }
}

Write-Host ""

# -----------------------------
# STEP 2: Old Report File
# -----------------------------

# Base path for historical weekly reports
$OldReportBase = "Q:\CORE\Transfers In\Reporting\Weekly Reports\Weekly Rollover Report 2026"

# Try to find the most recent xlsx under any dated subfolder (M-D-YYYY or MM-DD-YYYY)
$AssumedOldReport = $null
try {
    $AssumedOldReport = Get-ChildItem $OldReportBase -Recurse -File -Filter "*.xlsx" -ErrorAction SilentlyContinue |
        Where-Object { $_.DirectoryName -match '\d{1,2}-\d{1,2}-\d{4}$' } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1
} catch {
    # Q: drive may not be accessible on this machine
}

$OldReportFile = ""
$OldReportValid = $false
while (-not $OldReportValid) {
    Write-Host "Step 2: Last Week's Report (for carry-over data)" -ForegroundColor Yellow

    if ($AssumedOldReport) {
        Write-Host "Assumed file: $($AssumedOldReport.FullName)" -ForegroundColor Cyan
    } else {
        Write-Host "No file auto-detected under: $OldReportBase" -ForegroundColor Cyan
    }

    Write-Host "Press Enter to accept, paste a different path, or type SKIP:"
    $OldReportInput = Read-Host "Old report path"

    if ($OldReportInput.Trim().ToUpper() -eq "SKIP") {
        $OldReportFile = ""
        Write-Host "Skipping old report - carry-over data will not be included." -ForegroundColor Yellow
        $OldReportValid = $true
    } elseif ([string]::IsNullOrWhiteSpace($OldReportInput)) {
        if ($AssumedOldReport) {
            $OldReportFile = $AssumedOldReport.FullName
            Write-Host "File verified successfully!" -ForegroundColor Green
            $OldReportValid = $true
        } else {
            Write-Host ""
            Write-Host "ERROR: No file auto-detected. Enter a path or type SKIP." -ForegroundColor Red
            Write-Host ""
        }
    } else {
        $OldReportFile = $OldReportInput.Trim().Trim('"', "'")
        if (Test-Path $OldReportFile) {
            Write-Host "File verified successfully!" -ForegroundColor Green
            $OldReportValid = $true
        } else {
            Write-Host ""
            Write-Host "ERROR: File does not exist: $OldReportFile" -ForegroundColor Red
            Write-Host ""
        }
    }
}

Write-Host ""
Write-Host "=== Configuration Complete ===" -ForegroundColor Green
Write-Host ""

# -----------------------------
# Build the command
# -----------------------------
$OldReportArg = if ([string]::IsNullOrWhiteSpace($OldReportFile)) { "" } else { " -OldReportFile `"$OldReportFile`"" }
$Command = "powershell.exe -ExecutionPolicy Bypass -File `"$RunScriptPath`" " +
           "-RawInputFile `"$RawInputFile`"" + $OldReportArg

Write-Host "=== YOUR COMMAND ===" -ForegroundColor Green
Write-Host ""
Write-Host $Command -ForegroundColor White -BackgroundColor DarkBlue
Write-Host ""

# Save to clipboard
try {
    $Command | Set-Clipboard
    Write-Host "Command copied to clipboard!" -ForegroundColor Green
}
catch {
    Write-Host "Note: Command not copied to clipboard (manual copy required)" -ForegroundColor Yellow
}

Write-Host ""
Write-Host "=== VERIFICATION ===" -ForegroundColor Cyan

Write-Host "Raw input file exists: " -NoNewline
if (Test-Path $RawInputFile) {
    Write-Host "YES" -ForegroundColor Green
} else {
    Write-Host "NO - Check the path!" -ForegroundColor Red
}

Write-Host "Config file exists: " -NoNewline
if (Test-Path $ConfigPath) {
    Write-Host "YES" -ForegroundColor Green
} else {
    Write-Host "NO - Check the path!" -ForegroundColor Red
}

Write-Host "Old report file: " -NoNewline
if ([string]::IsNullOrWhiteSpace($OldReportFile)) {
    Write-Host "SKIPPED" -ForegroundColor Yellow
} elseif (Test-Path $OldReportFile) {
    Write-Host "YES" -ForegroundColor Green
} else {
    Write-Host "NO - Check the path!" -ForegroundColor Red
}

Write-Host ""
$RunNow = Read-Host "Run now? (Y/N)"

if ($RunNow -eq 'Y' -or $RunNow -eq 'y') {
    Write-Host ""
    Write-Host "=== Starting Processing ===" -ForegroundColor Green
    Write-Host ""

    if ([string]::IsNullOrWhiteSpace($OldReportFile)) {
        & $RunScriptPath -RawInputFile $RawInputFile
    } else {
        & $RunScriptPath -RawInputFile $RawInputFile -OldReportFile $OldReportFile
    }
} else {
    Write-Host ""
    Read-Host "Press Enter to exit"
}
