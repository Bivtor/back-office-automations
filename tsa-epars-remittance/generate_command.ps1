# TPA Remittance Processing - Command Generator
# This script helps users configure and run the remittance processor

Write-Host "=== TPA Remittance Processing - Command Generator ===" -ForegroundColor Green
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

# Get current username for Downloads path
$CurrentUser = $env:USERNAME

# Calculate today's ACH folder path
$Today = Get-Date
$Year = $Today.Year
$Month = $Today.ToString("MM")  # Two digits with leading zero
$Day = $Today.Day               # No leading zero for day

$ACHBase = "Q:\CORE\PURCHASE TRADING\ACH Remittance Detail"
$AssumedACHFolder = "$ACHBase\ACH-$Year\$Month - ACH\$Day"

# Question 1: Input_TPA location
Write-Host "Step 1: Input_TPA Location" -ForegroundColor Yellow
Write-Host "Default: Q:\CORE\PURCHASE TRADING\Input_TPA" -ForegroundColor Cyan
Write-Host "Press Enter to use default, or paste a different path:"
$InputTPAInput = Read-Host "Input_TPA path"

if ([string]::IsNullOrWhiteSpace($InputTPAInput)) {
    $InputTPA = "Q:\CORE\PURCHASE TRADING\Input_TPA"
} else {
    $InputTPA = $InputTPAInput.Trim().Trim('"', "'")
}

Write-Host ""

# Question 2: Downloads folder location
Write-Host "Step 2: EPARS/TSA Remittance Downloads Location" -ForegroundColor Yellow
Write-Host "Default: C:\Users\$CurrentUser\Downloads" -ForegroundColor Cyan
Write-Host "Press Enter to use default, or paste a different path:"
$DownloadsInput = Read-Host "Downloads path"

if ([string]::IsNullOrWhiteSpace($DownloadsInput)) {
    $DownloadsFolder = "C:\Users\$CurrentUser\Downloads"
} else {
    $DownloadsFolder = $DownloadsInput.Trim().Trim('"', "'")
}

Write-Host ""

# Question 3: Today's ACH folder with validation loop
$ACHFolderValid = $false
while (-not $ACHFolderValid) {
    Write-Host "Step 3: Today's ACH Folder Location" -ForegroundColor Yellow
    Write-Host "Assumed path: $AssumedACHFolder" -ForegroundColor Cyan
    Write-Host "Press Enter to accept, or paste a different path:"
    $ACHInput = Read-Host "ACH folder path"
    
    if ([string]::IsNullOrWhiteSpace($ACHInput)) {
        $ACHFolder = $AssumedACHFolder
    } else {
        $ACHFolder = $ACHInput.Trim().Trim('"', "'")
    }
    
    # Check if folder exists
    if (Test-Path $ACHFolder) {
        Write-Host "Folder verified successfully!" -ForegroundColor Green
        $ACHFolderValid = $true
    } else {
        Write-Host ""
        Write-Host "ERROR: Folder does not exist: $ACHFolder" -ForegroundColor Red
        Write-Host "Please verify the path and try again." -ForegroundColor Red
        Write-Host ""
    }
}

Write-Host ""

# Question 4: Morning Report Excel File with validation loop
$ExcelFileValid = $false
while (-not $ExcelFileValid) {
    Write-Host "Step 4: Morning Report Excel File Location" -ForegroundColor Yellow

    # Calculate today's Lockbox folder path
    $Today = Get-Date
    $Year = $Today.Year
    $Month = $Today.ToString("MM")
    $MonthName = $Today.ToString("MMMM")  # Full month name like "January"
    $Day = $Today.Day

    $LockboxBase = "Q:\CORE\PURCHASE TRADING\Lockbox Deposit Images"
    $YearFolder = "$LockboxBase\$Year Lockbox Deposit Images"
    $MonthFolder = "$YearFolder\$Month $MonthName Lockbox Deposit Images"

    # Try with leading zero first, then without
    $DayFolderWithZero = "$MonthFolder\$($Day.ToString('00'))"
    $DayFolderNoZero = "$MonthFolder\$Day"

    $DayFolder = $null
    if (Test-Path $DayFolderWithZero) {
        $DayFolder = $DayFolderWithZero
    } elseif (Test-Path $DayFolderNoZero) {
        $DayFolder = $DayFolderNoZero
    }

    $MorningReportPath = $null

    if ($DayFolder) {
        # Search for Morning Report file
        $AllFiles = Get-ChildItem -Path $DayFolder -Filter "*.xlsx" -File -ErrorAction SilentlyContinue
        
        foreach ($file in $AllFiles) {
            # Remove number prefix (everything before first space) and check if it matches
            $fileName = $file.Name
            if ($fileName -match '^\d+\s+(.+)$') {
                $nameWithoutPrefix = $matches[1]
                if ($nameWithoutPrefix -eq "ACHs Morning Report.xlsx") {
                    $MorningReportPath = $file.FullName
                    break
                }
            }
        }
    }

    if ($MorningReportPath) {
        Write-Host "Found: $MorningReportPath" -ForegroundColor Cyan
        Write-Host "Press Enter to use this file, or paste a different path:"
    } else {
        Write-Host "Could not auto-detect Morning Report file" -ForegroundColor Yellow
        Write-Host "Please paste the full path to the Morning Report Excel file:"
    }

    $ExcelInput = Read-Host "Excel file path"

    if ([string]::IsNullOrWhiteSpace($ExcelInput)) {
        if ($MorningReportPath) {
            $ExcelFile = $MorningReportPath
        } else {
            Write-Host ""
            Write-Host "ERROR: No file path provided and auto-detection failed" -ForegroundColor Red
            Write-Host "Please provide a valid path." -ForegroundColor Red
            Write-Host ""
            continue
        }
    } else {
        $ExcelFile = $ExcelInput.Trim().Trim('"', "'")
    }
    
    # Check if file exists
    if (Test-Path $ExcelFile) {
        Write-Host "File verified successfully!" -ForegroundColor Green
        $ExcelFileValid = $true
    } else {
        Write-Host ""
        Write-Host "ERROR: File does not exist: $ExcelFile" -ForegroundColor Red
        Write-Host "Please verify the path and try again." -ForegroundColor Red
        Write-Host ""
    }
}

Write-Host ""

# Question 5: PDF Password
Write-Host "Step 5: PDF Password" -ForegroundColor Yellow
Write-Host "Default: <REDACTED>" -ForegroundColor Cyan
Write-Host "Press Enter to use default, or type a different password:"
$PasswordInput = Read-Host "PDF password"

if ([string]::IsNullOrWhiteSpace($PasswordInput)) {
    $PDFPassword = "<REDACTED>"
} else {
    $PDFPassword = $PasswordInput.Trim()
}

Write-Host ""
Write-Host "=== Configuration Complete ===" -ForegroundColor Green
Write-Host ""

# Build the command
$Command = "powershell.exe -ExecutionPolicy Bypass -File `"$RunScriptPath`" -ExcelFile `"$ExcelFile`" -InputTPA `"$InputTPA`" -Downloads `"$DownloadsFolder`" -ACHFolder `"$ACHFolder`" -PDFPassword `"$PDFPassword`""
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
Write-Host "Excel file exists: " -NoNewline
if (Test-Path $ExcelFile) {
    Write-Host "YES" -ForegroundColor Green
} else {
    Write-Host "NO - Check the path!" -ForegroundColor Red
}

Write-Host "Input_TPA exists: " -NoNewline
if (Test-Path $InputTPA) {
    Write-Host "YES" -ForegroundColor Green
} else {
    Write-Host "NO - Check the path!" -ForegroundColor Red
}

Write-Host "Downloads folder exists: " -NoNewline
if (Test-Path $DownloadsFolder) {
    Write-Host "YES" -ForegroundColor Green
} else {
    Write-Host "NO - Check the path!" -ForegroundColor Red
}

Write-Host "ACH folder exists: " -NoNewline
if (Test-Path $ACHFolder) {
    Write-Host "YES" -ForegroundColor Green
} else {
    Write-Host "NO - Check the path!" -ForegroundColor Red
}
Write-Host ""
Write-Host "=== INSTRUCTIONS ===" -ForegroundColor Yellow
Write-Host "1. Copy the blue command above (already in clipboard)"
Write-Host "2. Press Windows Key + R"
Write-Host "3. Type 'cmd' and press Enter"
Write-Host "4. Paste the command and press Enter"
Write-Host ""
Write-Host "OR simply close this window and the script will run automatically..."
Write-Host ""

$RunNow = Read-Host "Run now? (Y/N)"

if ($RunNow -eq 'Y' -or $RunNow -eq 'y') {
    Write-Host ""
    Write-Host "=== Starting Processing ===" -ForegroundColor Green
    Write-Host ""
    
    & $RunScriptPath -ExcelFile $ExcelFile -InputTPA $InputTPA -Downloads $DownloadsFolder -ACHFolder $ACHFolder -PDFPassword $PDFPassword} else {
    Write-Host ""
    Read-Host "Press Enter to exit"
}
