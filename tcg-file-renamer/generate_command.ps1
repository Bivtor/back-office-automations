# TCG File Renamer - Command Generator
# This script helps users create the correct command to run

Write-Host "=== TCG File Renamer - Command Generator ===" -ForegroundColor Green
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

# Question 1: Excel File Location with auto-detection and validation loop
$ExcelFileValid = $false
while (-not $ExcelFileValid) {
    Write-Host "Step 1: Morning Report Excel File Location" -ForegroundColor Yellow

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
            $ExcelPath = $MorningReportPath
        } else {
            Write-Host ""
            Write-Host "ERROR: No file path provided and auto-detection failed" -ForegroundColor Red
            Write-Host "Please provide a valid path." -ForegroundColor Red
            Write-Host ""
            continue
        }
    } else {
        $ExcelPath = $ExcelInput.Trim().Trim('"', "'")
    }
    
    # Check if file exists
    if (Test-Path $ExcelPath) {
        Write-Host "File verified successfully!" -ForegroundColor Green
        $ExcelFileValid = $true
    } else {
        Write-Host ""
        Write-Host "ERROR: File does not exist: $ExcelPath" -ForegroundColor Red
        Write-Host "Please verify the path and try again." -ForegroundColor Red
        Write-Host ""
    }
}

Write-Host ""

# Question 2: Folder Location (Input_TPA) with validation loop
$FolderValid = $false
while (-not $FolderValid) {
    Write-Host "Step 2: TCG Files Folder Location (Input_TPA)" -ForegroundColor Yellow
    Write-Host "Default: Q:\CORE\PURCHASE TRADING\Input_TPA" -ForegroundColor Cyan
    Write-Host "Press Enter to use default, or paste a different path:"
    $FolderInput = Read-Host "Files folder path"

    if ([string]::IsNullOrWhiteSpace($FolderInput)) {
        $FolderPath = "Q:\CORE\PURCHASE TRADING\Input_TPA"
    } else {
        $FolderPath = $FolderInput.Trim().Trim('"', "'")
    }
    
    # Check if folder exists
    if (Test-Path $FolderPath) {
        Write-Host "Folder verified successfully!" -ForegroundColor Green
        $FolderValid = $true
    } else {
        Write-Host ""
        Write-Host "ERROR: Folder does not exist: $FolderPath" -ForegroundColor Red
        Write-Host "Please verify the path and try again." -ForegroundColor Red
        Write-Host ""
    }
}

Write-Host ""

# Question 3: Downloads folder location with validation loop
$DownloadsValid = $false
while (-not $DownloadsValid) {
    Write-Host "Step 3: Downloads Folder Location (for log file)" -ForegroundColor Yellow
    Write-Host "Default: C:\Users\$CurrentUser\Downloads" -ForegroundColor Cyan
    Write-Host "Press Enter to use default, or paste a different path:"
    $DownloadsInput = Read-Host "Downloads path"

    if ([string]::IsNullOrWhiteSpace($DownloadsInput)) {
        $DownloadsFolder = "C:\Users\$CurrentUser\Downloads"
    } else {
        $DownloadsFolder = $DownloadsInput.Trim().Trim('"', "'")
    }
    
    # Check if folder exists
    if (Test-Path $DownloadsFolder) {
        Write-Host "Folder verified successfully!" -ForegroundColor Green
        $DownloadsValid = $true
    } else {
        Write-Host ""
        Write-Host "ERROR: Folder does not exist: $DownloadsFolder" -ForegroundColor Red
        Write-Host "Please verify the path and try again." -ForegroundColor Red
        Write-Host ""
    }
}

Write-Host ""

# Question 4: File Name Template
# Get today's date in M-D format (changed from MM-DD)
$TodayDate = Get-Date -Format "M-d"
$DefaultTemplate = "JNT [EMPLOYER_NAME] $TodayDate [AMOUNT]"

Write-Host "Step 4: File Name Template" -ForegroundColor Yellow
Write-Host "Default template: $DefaultTemplate" -ForegroundColor Cyan
Write-Host "[EMPLOYER_NAME] and [AMOUNT] will be replaced with actual values from Excel"
Write-Host ""
Write-Host "Press Enter to use default, or type your custom template:"
$TemplateInput = Read-Host "File name template"

if ([string]::IsNullOrWhiteSpace($TemplateInput)) {
    $Template = $DefaultTemplate
} else {
    $Template = $TemplateInput.Trim()
}

Write-Host ""
Write-Host "=== Configuration Complete ===" -ForegroundColor Green
Write-Host ""

# Build the command
$Command = "powershell.exe -ExecutionPolicy Bypass -File `"$RunScriptPath`" -ExcelFile `"$ExcelPath`" -RemittanceFolder `"$FolderPath`" -Downloads `"$DownloadsFolder`" -FileNameTemplate `"$Template`""

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
if (Test-Path $ExcelPath) {
    Write-Host "YES" -ForegroundColor Green
} else {
    Write-Host "NO - Check the path!" -ForegroundColor Red
}

Write-Host "Folder exists: " -NoNewline
if (Test-Path $FolderPath) {
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
    
    & $RunScriptPath -ExcelFile $ExcelPath -RemittanceFolder $FolderPath -Downloads $DownloadsFolder -FileNameTemplate $Template
} else {
    Write-Host ""
    Read-Host "Press Enter to exit"
}
