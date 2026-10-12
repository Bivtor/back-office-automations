# TPA Remittance Processing - Main Script
# Processes EPARS/TSA remittance files with PDF extraction and matching

param(
    [Parameter(Mandatory=$true)]
    [string]$ExcelFile,
    
    [Parameter(Mandatory=$true)]
    [string]$InputTPA,
    
    [Parameter(Mandatory=$true)]
    [string]$Downloads,
    
    [Parameter(Mandatory=$true)]
    [string]$ACHFolder,
    
    [Parameter(Mandatory=$true)]
    [string]$PDFPassword
)

# Clean up parameters
$ExcelFile = $ExcelFile.Trim('"', ' ')
$InputTPA = $InputTPA.Trim('"', ' ')
$Downloads = $Downloads.Trim('"', ' ')
$ACHFolder = $ACHFolder.Trim('"', ' ')

# Get script directory
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

# Generate log file name with timestamp
$LogFile = Join-Path $Downloads "RemittanceProcessing_$(Get-Date -Format 'MM-dd-yyyy_HH-mm').log"

# ================= LOGGING =================

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARNING','ERROR','MATCH')]
        [string]$Level = 'INFO'
    )

    $Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $Line = "$Timestamp - $Level - $Message"

    switch ($Level) {
        'INFO'    { Write-Host $Line -ForegroundColor White }
        'WARNING' { Write-Host $Line -ForegroundColor Yellow }
        'ERROR'   { Write-Host $Line -ForegroundColor Red }
        'MATCH'   { Write-Host $Line -ForegroundColor Green }
    }

    Add-Content -Path $LogFile -Value $Line
}

Write-Log "========================================"
Write-Log "TPA Remittance Processing Started"
Write-Log "========================================"
Write-Log "Excel File: $ExcelFile"
Write-Log "Input TPA: $InputTPA"
Write-Log "Downloads: $Downloads"
Write-Log "ACH Folder: $ACHFolder"
Write-Log "PDF Password: [REDACTED]"
Write-Log "Log File: $LogFile"
Write-Log ""

# ================= CONFIGURATION =================

# Load JSON configuration file
$EPARSNamesPath = Join-Path $ScriptDir "EPARS_NAMES.json"
if (-not (Test-Path $EPARSNamesPath)) {
    Write-Log "ERROR: EPARS_NAMES.json not found at $EPARSNamesPath" -Level ERROR
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

try {
    $Config = Get-Content $EPARSNamesPath -Raw | ConvertFrom-Json
    
    # Validate JSON structure
    if (-not $Config.remittance_mappings) {
        throw "JSON missing 'remittance_mappings' property"
    }
    
    # Create dictionary mapping excel_name -> file_name
    $NameMappingDict = @{}
    foreach ($mapping in $Config.remittance_mappings) {
        if (-not $mapping.excel_name -or -not $mapping.file_name) {
            throw "Invalid mapping found - missing excel_name or file_name"
        }
        $NameMappingDict[$mapping.excel_name.ToLower()] = $mapping.file_name
    }
    
    Write-Log "Loaded EPARS_NAMES.json configuration successfully"
    Write-Log "Found $($NameMappingDict.Count) employer mappings"
    
} catch {
    Write-Log "ERROR: Failed to parse EPARS_NAMES.json: $($_.Exception.Message)" -Level ERROR
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Log ""




# ================= HELPER FUNCTIONS =================


function Read-PasswordProtectedCSV {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FilePath,

        [Parameter(Mandatory=$true)]
        [string]$Password
    )

    # Placeholder implementation
    # Replace with actual decryption logic if needed
    # Currently assumes file is readable as plain text

    return Get-Content -Path $FilePath
}


function Get-CSVColumnSum {
    param(
        [Parameter(Mandatory=$true)]
        [string]$CSVPath,

        [Parameter(Mandatory=$true)]
        [string]$Password,

        [Parameter(Mandatory=$true)]
        [bool]$IsTSA
    )

    if (-not (Test-Path $CSVPath)) {
        throw "CSV file not found: $CSVPath"
    }

    $Lines = Get-Content -Path $CSVPath
	$Total = 0.0

	# Determine column index
	$ColumnIndex = if ($IsTSA) { 7 } else { 1 }

	# Skip header row for TSA
	$StartIndex = if ($IsTSA) { 0 } else { 1 }

    for ($i = $StartIndex; $i -lt $Lines.Count; $i++) {

		$Line = $Lines[$i]

		if ([string]::IsNullOrWhiteSpace($Line)) {
			continue
		}

		$Columns = $Line -split ','

		if ($Columns.Count -le $ColumnIndex) {
			continue
		}

		$RawValue = $Columns[$ColumnIndex] -replace '[$,"]', ''
		$ParsedValue = 0.0

		if ([double]::TryParse($RawValue, [ref]$ParsedValue)) {
			$Total += $ParsedValue
		}
	}


    return [Math]::Round($Total, 2)
}



function Find-CSVFileInFolder {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FolderPath
    )

    $CSVFiles = Get-ChildItem -Path $FolderPath -File |
        Where-Object { $_.Name -like '*_CSV.txt' }

    if ($CSVFiles.Count -eq 0) {
        return $null
    }

    if ($CSVFiles.Count -gt 1) {
        Write-Log "  Multiple _CSV.txt files found - using first" -Level WARNING
    }

    return $CSVFiles | Select-Object -First 1
}

function Extract-PasswordProtectedZip {
    param(
        [string]$ZipPath,
        [string]$DestinationPath,
        [string]$Password
    )

    $7zipPaths = @(
        "C:\Program Files\7-Zip\7z.exe",
        "C:\Program Files (x86)\7-Zip\7z.exe"
    )

    $7zipExe = $7zipPaths | Where-Object { Test-Path $_ } | Select-Object -First 1

    if (-not $7zipExe) {
        throw "7-Zip not found. Please install 7-Zip."
    }

    $arguments = @(
        "x"
        "`"$ZipPath`""
        "-o`"$DestinationPath`""
        "-p`"$Password`""
        "-y"
        "-bd"   # disable progress indicator
        "-bb0"  # suppress output
    )

	$nullOut = "$env:TEMP\7zip_stdout.txt"
	$nullErr = "$env:TEMP\7zip_stderr.txt"

	$process = Start-Process `
		-FilePath $7zipExe `
		-ArgumentList $arguments `
		-Wait `
		-NoNewWindow `
		-RedirectStandardOutput $nullOut `
		-RedirectStandardError  $nullErr `
		-PassThru




    if ($process.ExitCode -ne 0) {
        throw "7-Zip extraction failed with exit code: $($process.ExitCode)"
    }
}

# Step 3 helpers

function Find-PDFInFolder {
    param(
        [Parameter(Mandatory=$true)]
        [string]$FolderPath
    )

    Get-ChildItem -Path $FolderPath -Recurse -File |
        Where-Object { $_.Extension -eq '.pdf' } |
        Select-Object -First 1
}

function Get-PDFFirstPageText {
    param(
        [Parameter(Mandatory=$true)]
        [string]$PdfPath,

        [Parameter(Mandatory=$true)]
        [string]$PdfToTextExe
    )

    $TempTxt = [System.IO.Path]::GetTempFileName()

    try {
        & $PdfToTextExe `
            -f 1 `
            -l 1 `
            -layout `
            "$PdfPath" `
            "$TempTxt" | Out-Null

        Get-Content -Path $TempTxt -Raw
    }
    finally {
        if (Test-Path $TempTxt) {
            Remove-Item $TempTxt -Force
        }
    }
}


function Normalize-PDFText {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Text
    )

    $Normalized = $Text `
        -replace "`r", "" `
        -replace "[ ]{2,}", " " `
        -replace "`n{2,}", "`n"

    return $Normalized.Trim()
}



function Parse-TSAFirstPage {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Text
    )

    # Placeholder region selection logic goes here
    $TargetRegion = $Text

    # Example regex only
    if ($TargetRegion -match "Batch\s+ID\s*:\s*(\d+)") {
        return $Matches[1]
    }

    throw "Unable to parse TSA identifier from PDF"
}


function Parse-EPARSFirstPage {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Text
    )

    # Placeholder region selection logic goes here
    $TargetRegion = $Text

    # Example regex only
    if ($TargetRegion -match "Control\s+Number\s*:\s*(\d+)") {
        return $Matches[1]
    }

    throw "Unable to parse EPARS identifier from PDF"
}


function Rename-ExtractedFolderAndFiles {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ExtractPath,

        [Parameter(Mandatory=$true)]
        [string]$NewBaseName
    )

    $Parent = Split-Path $ExtractPath -Parent
    $NewFolderPath = Join-Path $Parent $NewBaseName

    if (Test-Path $NewFolderPath) {
        throw "Target folder already exists: $NewFolderPath"
    }

    Get-ChildItem -Path $ExtractPath -File -Recurse | ForEach-Object {
        $NewName = "$NewBaseName$($_.Extension)"
        Rename-Item $_.FullName $NewName -Force
    }

    Rename-Item $ExtractPath $NewFolderPath -Force

    return $NewFolderPath
}

function Move-ProcessedFiles {
    param(
        [Parameter(Mandatory=$true)]
        [string]$PDFPath,
        
        [Parameter(Mandatory=$true)]
        [string]$CSVPath,
        
        [Parameter(Mandatory=$true)]
        [string]$ACHFolder,
        
        [Parameter(Mandatory=$true)]
        [string]$InputTPAFolder
    )
    
    try {
        # Move PDF to ACH Folder
        $PDFFileName = Split-Path $PDFPath -Leaf
        $PDFDestination = Join-Path $ACHFolder $PDFFileName
        Move-Item -Path $PDFPath -Destination $PDFDestination -Force
        Write-Log "    Moved PDF to: $PDFDestination" -Level MATCH
        
        # Move CSV/TXT to Input TPA Folder
        $CSVFileName = Split-Path $CSVPath -Leaf
        $CSVDestination = Join-Path $InputTPAFolder $CSVFileName
        Move-Item -Path $CSVPath -Destination $CSVDestination -Force
        Write-Log "    Moved CSV to: $CSVDestination" -Level MATCH
        
        return $true
    }
    catch {
        Write-Log "    Error moving files: $($_.Exception.Message)" -Level ERROR
        return $false
    }
}

# [AMENDMENT 3] Cleans up the extracted temp folder and original ZIP after successful processing
function Remove-ProcessedZipAndFolder {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ZipPath,

        [Parameter(Mandatory=$true)]
        [string]$ExtractPath
    )

    $cleanedUp = $true

    # Remove extracted temp folder
    if (Test-Path $ExtractPath) {
        try {
            Remove-Item -Path $ExtractPath -Recurse -Force
            Write-Log "    Deleted extract folder: $ExtractPath" -Level INFO
        }
        catch {
            Write-Log "    WARNING: Could not delete extract folder '$ExtractPath': $($_.Exception.Message)" -Level WARNING
            $cleanedUp = $false
        }
    }

    # Remove original ZIP from Downloads
    if (Test-Path $ZipPath) {
        try {
            Remove-Item -Path $ZipPath -Force
            Write-Log "    Deleted source ZIP: $ZipPath" -Level INFO
        }
        catch {
            Write-Log "    WARNING: Could not delete ZIP '$ZipPath': $($_.Exception.Message)" -Level WARNING
            $cleanedUp = $false
        }
    }

    return $cleanedUp
}


# ================= STEP 1 - EXCEL EXTRACTION =================

function Extract-RemittanceData {
    param(
        [string]$ExcelPath,
        [hashtable]$NameMapping
    )

    if (-not (Test-Path $ExcelPath)) {
        throw "Excel file not found: $ExcelPath"
    }

    $Excel = $null
    $Workbook = $null
    
    try {
        Write-Log "Attempting to open Excel..."
        
        # Retry logic for Excel COM object
        $MaxRetries = 5
        $RetryCount = 0
        $Success = $false
        
        while (-not $Success -and $RetryCount -lt $MaxRetries) {
            try {
                Write-Log "Creating Excel COM object (attempt $($RetryCount + 1) of $MaxRetries)..."
                $Excel = New-Object -ComObject Excel.Application
                Start-Sleep -Seconds 2
                
                $Excel.Visible = $false
                $Excel.DisplayAlerts = $false
                $Excel.AskToUpdateLinks = $false
                $Excel.AutomationSecurity = 3
                $Excel.ScreenUpdating = $false
                
                Write-Log "Opening workbook: $ExcelPath"
                Write-Log "Please wait, this may take a moment on slower systems..."
                
                $Workbook = $Excel.Workbooks.Open($ExcelPath, 0, $true)
                Start-Sleep -Seconds 3
                
                # Verify workbook is ready
                $null = $Workbook.Worksheets.Count
                
                $Success = $true
                Write-Log "Excel workbook opened successfully"
            }
            catch {
                $RetryCount++
                Write-Log "Attempt $RetryCount failed: $($_.Exception.Message)" -Level WARNING
                
                if ($RetryCount -lt $MaxRetries) {
                    Write-Log "Retrying in 5 seconds..." -Level WARNING
                    Start-Sleep -Seconds 5
                    
                    # Cleanup failed attempt
                    if ($Workbook) {
                        try { $Workbook.Close($false) } catch {}
                        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook) | Out-Null } catch {}
                    }
                    if ($Excel) {
                        try { $Excel.Quit() } catch {}
                        try { [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null } catch {}
                    }
                    
                    [System.GC]::Collect()
                    [System.GC]::WaitForPendingFinalizers()
                    Start-Sleep -Seconds 2
                    
                    $Excel = $null
                    $Workbook = $null
                } else {
                    throw "Failed to open Excel after $MaxRetries attempts. Please close all Excel windows and try again."
                }
            }
        }

        $Worksheet = $Workbook.Worksheets.Item(1)
        Write-Log "Reading data from worksheet..."

        $Results = @()
        $AmountTracker = @{}
        $MaxRow = $Worksheet.UsedRange.Rows.Count
        Write-Log "Scanning $MaxRow rows for remittance data..."

        # TODO: Adjust these column numbers based on your actual Excel structure
        $EmployerNameColumn = 2  # Column B
        $AmountColumn = 9        # Column I
        
        for ($Row = 2; $Row -le $MaxRow; $Row++) {
            
            # Read employer name (full text from Excel)
            $EmployerNameFull = $Worksheet.Cells.Item($Row, $EmployerNameColumn).Text
            
            # Skip empty rows
            if ([string]::IsNullOrWhiteSpace($EmployerNameFull)) {
                continue
            }
            
            $EmployerNameFull = $EmployerNameFull.Trim()
            
            # Extract name up to first comma for matching
            $EmployerName = $EmployerNameFull
            $CommaIndex = $EmployerNameFull.IndexOf(",")
            if ($CommaIndex -gt 0) {
                $EmployerName = $EmployerNameFull.Substring(0, $CommaIndex).Trim()
            }
            
            # Check if this employer is in our mapping (case-insensitive)
            $EmployerNameLower = $EmployerName.ToLower()
			# Write-Log "Row $Row - Checking: "$EmployerNameLower"" # DEBUG LOG

            if (-not $NameMapping.ContainsKey($EmployerNameLower)) {
                continue  # Skip employers not in our mapping
            }
            
            # Get the file name from our mapping
            $FileNamePrefix = $NameMapping[$EmployerNameLower]
			# Write-Log "  -> MATCHED to: $FileNamePrefix" -Level MATCH # DEBUG LOG
            
            # Read amount from merged cell (try both columns)
            # When cells are merged, Excel stores the value in the top-left cell
            $AmountText = $Worksheet.Cells.Item($Row - 1, $AmountColumn).Text
            if (-not $AmountText -or $AmountText -eq "") {
                $AmountText = $Worksheet.Cells.Item($Row - 1, $AmountColumn + 1).Text
            }
			
            $AmountText = $AmountText -replace "[$,]", ""  # Remove $ and commas
            if (-not $AmountText -or $AmountText -eq "") { $AmountText = "0.00" }
            
			# Write-log "Amount text is - $AmountText" # DEBUG LOG
            # Parse and format amount
            $Amount = 0.0
            if ([double]::TryParse($AmountText, [ref]$Amount)) {
                
                # Format amount - cleaning up trailing .00 when amounts are whole number
                $FormattedAmount = if ($Amount % 1 -eq 0) { 
                    $Amount.ToString("0") 
                } else { 
                    $Amount.ToString("0.00") 
                }
				
			} else {
				Write-log "Could Not parse amount"
			}
                
                # Create target object
                $target = [PSCustomObject]@{
                    Row = $Row
                    EmployerName = $EmployerName
                    FileNamePrefix = $FileNamePrefix
                    Amount = [double]$Amount
                    FormattedAmount = $FormattedAmount
                    MatchedFolder = $null
                    Status = "Pending"
                }
                
                $Results += $target
            
			
        }
		
		# Check for duplicate amounts and remove them
        Write-Log ""
        Write-Log "Checking for duplicate amounts..."
        $AmountGroups = $Results | Group-Object -Property Amount | Where-Object { $_.Count -gt 1 }
        
        if ($AmountGroups) {
            Write-Log "WARNING: Found duplicate amounts - removing all duplicates from results" -Level WARNING
            foreach ($group in $AmountGroups) {
                Write-Log "  Duplicate amount: $($group.Name)" -Level WARNING
                foreach ($item in $group.Group) {
                    Write-Log "    - Row $($item.Row): $($item.EmployerName) - $($item.FormattedAmount)" -Level WARNING
                }
            }
            
            # Get all duplicate amounts
            $DuplicateAmounts = $AmountGroups | ForEach-Object { [double]$_.Name }
            
            # Remove all entries with duplicate amounts
            $OriginalCount = $Results.Count
            $Results = $Results | Where-Object { $DuplicateAmounts -notcontains $_.Amount }
            $RemovedCount = $OriginalCount - $Results.Count - 1
            
            Write-Log "Removed $RemovedCount entries due to duplicate amounts" -Level WARNING
        } else {
            Write-Log "No duplicate amounts found"
        }
        
        Write-Log ""
        Write-Log "Found $($Results.Count) remittance entries"
        
        # Log all entries
        Write-Log ""
        Write-Log "All Entries:" -Level MATCH
        foreach ($entry in $Results) {
            Write-Log "  $($entry.FileNamePrefix) - $($entry.FormattedAmount)" -Level MATCH
        }

        return $Results
    }
    finally {
        # Proper COM object cleanup
        if ($Workbook) {
            $Workbook.Close($false)
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Workbook) | Out-Null
        }
        if ($Excel) {
            $Excel.Quit()
            [System.Runtime.Interopservices.Marshal]::ReleaseComObject($Excel) | Out-Null
        }
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}


# ================= STEP 2 - FOLDER MATCHING =================

function Match-DownloadedFolders {
    param(
        [Parameter(Mandatory=$true)]
        [string]$DownloadsPath,

        [Parameter(Mandatory=$true)]
        [array]$RemittanceTargets,

        [Parameter(Mandatory=$true)]
        [string]$Password
    )

    if (-not (Test-Path $DownloadsPath)) {
        throw "Downloads directory not found: $DownloadsPath"
    }

    Write-Log "Scanning Downloads directory for ZIP files downloaded TODAY..."

    # [AMENDMENT 2] Filter to today's ZIPs only, sorted newest -> oldest
    $Today = (Get-Date).Date
    $ZipFiles = Get-ChildItem -Path $DownloadsPath -File -Filter '*.zip' |
        Where-Object { $_.LastWriteTime.Date -eq $Today } |
        Sort-Object -Property LastWriteTime -Descending

    if ($ZipFiles.Count -eq 0) {
        Write-Log "No ZIP files downloaded today found in: $DownloadsPath" -Level WARNING
    } else {
        Write-Log "Found $($ZipFiles.Count) ZIP file(s) from today (processing newest first)"
    }

    # Cache: zip_full_path -> calculated amount
    $Results = @()

    foreach ($Zip in $ZipFiles) {
		
		# Determine TSA vs EPARS based on ZIP name
		$IsTSA = $Zip.BaseName.Substring(0, [Math]::Min(5, $Zip.BaseName.Length)) -match '^EPARS'


        Write-Log ""
        Write-Log "Processing ZIP: $($Zip.Name)"

		$ExtractPath = Join-Path $DownloadsPath "_extract_$($Zip.BaseName)"

		try {
			# Ensure clean extraction path
			if (Test-Path $ExtractPath) {
				Remove-Item -Path $ExtractPath -Recurse -Force
			}

			New-Item -Path $ExtractPath -ItemType Directory | Out-Null

			Write-Log "  Extracting ZIP to: $ExtractPath"
			Extract-PasswordProtectedZip -ZipPath $Zip.FullName -DestinationPath $ExtractPath -Password $Password
			$CSVFile = Find-CSVFileInFolder -FolderPath $ExtractPath

			if (-not $CSVFile) {
				Write-Log "  No _CSV.txt file found in ZIP - skipping" -Level WARNING
				continue
			}

			Write-Log "  Reading CSV file: $($CSVFile.Name)"

			$ZipAmount = Get-CSVColumnSum `
				-CSVPath $CSVFile.FullName `
				-Password $Password `
				-IsTSA $IsTSA


			$result = [PSCustomObject]@{
				ZipPath       = $Zip.FullName
				ExtractPath   = $ExtractPath
				Amount        = [double]$ZipAmount
				IsTSA         = $IsTSA
				MatchedTarget = $null
				PDF_Rename    = $null
				CSV_Rename    = $null
			}

			$Results += $result

			Write-Log "  Calculated ZIP amount: $ZipAmount"
		}
		catch {
			Write-Log "  Error processing ZIP $($Zip.Name): $($_.Exception.Message)" -Level ERROR
			continue
		}

		# ---------- MATCHING ----------
		$MatchedTarget = $RemittanceTargets | Where-Object {
			$_.Status -ne 'Matched' -and
			[Math]::Abs($_.Amount - $ZipAmount) -lt 0.01
		} | Select-Object -First 1

		if ($MatchedTarget) {
			$MatchedTarget.Status = 'Matched'
			$MatchedTarget.MatchedFolder = $Zip.FullName
			$result.MatchedTarget = $MatchedTarget

			Write-Log "MATCH: $($MatchedTarget.FileNamePrefix) -> $($Zip.Name)" -Level MATCH
		}
		else {
			Write-Log "  No matching remittance target for amount $ZipAmount" -Level WARNING
		}
	} # end FOR 

    # ---------- SUMMARY ----------
    Write-Log ""
    Write-Log "Step 2 Summary:"
    foreach ($target in $RemittanceTargets) {
        Write-Log "  $($target.FileNamePrefix) - $($target.FormattedAmount) - $($target.Status)"
    }

    # [AMENDMENT 1] Verbose unmatched warning with mandatory acknowledgement
    $UnmatchedTargets = $RemittanceTargets | Where-Object { $_.Status -ne 'Matched' }

    if ($UnmatchedTargets) {
        Write-Host ""
        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red
        Write-Host "!!                                                     !!" -ForegroundColor Red
        Write-Host "!!              *** UNMATCHED FILES ***                !!" -ForegroundColor Red
        Write-Host "!!                                                     !!" -ForegroundColor Red
        Write-Host "!!  The following remittances from Excel did NOT have  !!" -ForegroundColor Red
        Write-Host "!!  a matching ZIP file in the Downloads folder:       !!" -ForegroundColor Red
        Write-Host "!!                                                     !!" -ForegroundColor Red
        foreach ($unmatched in $UnmatchedTargets) {
            $line = "!!    - $($unmatched.FileNamePrefix) ($($unmatched.FormattedAmount))"
            # Pad to fixed width so the banner stays aligned
            $padded = $line.PadRight(55) + "!!"
            Write-Host $padded -ForegroundColor Red
        }
        Write-Host "!!                                                     !!" -ForegroundColor Red
        Write-Host "!!  Please confirm you are aware these were MISSED     !!" -ForegroundColor Red
        Write-Host "!!  before processing continues.                       !!" -ForegroundColor Red
        Write-Host "!!                                                     !!" -ForegroundColor Red
        Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!" -ForegroundColor Red
        Write-Host ""

        # Log the unmatched entries
        foreach ($unmatched in $UnmatchedTargets) {
            Write-Log "UNMATCHED: $($unmatched.FileNamePrefix) - $($unmatched.FormattedAmount)" -Level ERROR
        }

        Write-Host "Press Enter to acknowledge these missed files and continue..." -ForegroundColor Red
        Read-Host
        Write-Log "User acknowledged $($UnmatchedTargets.Count) unmatched remittance(s) and chose to continue"
    }

    return $Results
}

# ================= STEP 3 HELPER - EXTRACT AND RENAME =================

function Get-PrecomputedNames {
    param(
        [Parameter(Mandatory=$true)]
        [PSCustomObject]$ZipResult
    )

    $Today = Get-Date

    # Build date as M-d (no leading zeros)
    $DateString = "$($Today.Month)-$($Today.Day)"
    
    # Format amount
    if ($ZipResult.Amount % 1 -eq 0) {
		# Whole number → no decimals
		$AmountString = [int]$ZipResult.Amount
	}
	else {
		# Has cents → keep two decimals
		$AmountString = $ZipResult.Amount.ToString("F2")
	}

    if ($ZipResult.IsTSA) {
        # TSA: Both PDF and CSV get same name
        $BaseName = "TSA $DateString $AmountString"
        return @{
            PDFName = $BaseName
            CSVName = $BaseName
        }
    }
    else {
        # EPARS
        $MatchedPrefix = $ZipResult.MatchedTarget.FileNamePrefix
        $PDFName = "$MatchedPrefix $DateString $AmountString"
        $CSVName = "EPARS 403(b) EE $MatchedPrefix $DateString $AmountString"
        
        return @{
            PDFName = $PDFName
            CSVName = $CSVName
        }
    }
}




# ================= MAIN EXECUTION =================

# Check if Excel is running
$ExcelProcesses = Get-Process excel -ErrorAction SilentlyContinue
# if ($ExcelProcesses) {
#    Write-Log "WARNING: Excel is currently running. This may cause issues." -Level WARNING
#    Write-Log "Please close all Excel windows and press Enter to continue, or Ctrl+C to cancel." -Level WARNING
#    Read-Host
# }

try {
    # Step 1: Extract data from Excel
    Write-Log "========================================"
    Write-Log "STEP 1: Reading Excel Data"
    Write-Log "========================================"
    Write-Log ""
    
    $RemittanceTargets = Extract-RemittanceData -ExcelPath $ExcelFile -NameMapping $NameMappingDict
    
    if ($RemittanceTargets.Count -eq 0) {
        Write-Log "ERROR: No remittance entries found in Excel file" -Level ERROR
        Write-Host ""
        Read-Host "Press Enter to exit"
        exit 1
    }
    
	
	# Step 2: Scan downloaded folders and match by amount
    Write-Log "========================================"
    Write-Log "STEP 2: Scanning Downloads and Matching"
    Write-Log "========================================"
    Write-Log ""
    
    $ProcessedZips = Match-DownloadedFolders -DownloadsPath $Downloads -RemittanceTargets $RemittanceTargets -Password $PDFPassword

	$MatchedTargets = $RemittanceTargets | Where-Object Status -eq 'Matched'

	# Safety Check
	if (-not $MatchedTargets) {
		Write-Log "No matched remittances found - Step 3 skipped" -Level WARNING
		return
	}
	
	Write-Log "========================================"
	Write-Log "STEP 3: Reviewing PDF Data and Renaming"
	Write-Log "========================================"
	Write-Log ""


	foreach ($ZipResult in $ProcessedZips | Where-Object { $_.MatchedTarget }) {

		Write-Log "Processing matched ZIP: $($ZipResult.ZipPath)"

		$PdfFile = Find-PDFInFolder -FolderPath $ZipResult.ExtractPath
		$CSVFile = Find-CSVFileInFolder -FolderPath $ZipResult.ExtractPath

		if (-not $PdfFile) {
			Write-Log "  No PDF found in extracted folder" -Level WARNING
			continue
		}

		if (-not $CSVFile) {
			Write-Log "  No CSV found in extracted folder" -Level WARNING
			continue
		}

		Write-Log "  Found PDF: $($PdfFile.Name)"
		Write-Log "  Found CSV: $($CSVFile.Name)"

		# Get precomputed names
		$PrecomputedNames = Get-PrecomputedNames -ZipResult $ZipResult
		
		try {
			# Open PDF for user review
			Write-Log "  Opening PDF for review..."
			Start-Process $PdfFile.FullName
			Start-Sleep -Milliseconds 500

			# Prompt for PDF name with precomputed value shown
			Write-Host ""
			Write-Host "=== PDF Rename ===" -ForegroundColor Cyan
			
			# Copy suggestion to clipboard
			$PrecomputedNames.PDFName | Set-Clipboard

			Write-Host "Suggested (copied to clipboard):" -ForegroundColor Yellow
			Write-Host "  $($PrecomputedNames.PDFName)" -ForegroundColor Yellow
			
			$PDFResponse = Read-Host "Enter name (Enter = accept, Ctrl+V = paste & edit)"
			
			if ([string]::IsNullOrWhiteSpace($PDFResponse)) {
				$ZipResult.PDF_Rename = $PrecomputedNames.PDFName
			}
			else {
				$ZipResult.PDF_Rename = $PDFResponse
			}

			Write-Log "  PDF will be renamed to: $($ZipResult.PDF_Rename)"

			# Prompt for CSV name with precomputed value shown
			Write-Host ""
			Write-Host "=== CSV Rename ===" -ForegroundColor Cyan
			
			# Copy suggestion to clipboard
			$PrecomputedNames.CSVName | Set-Clipboard
			Write-Host "Suggested (copied to clipboard):" -ForegroundColor Yellow
			Write-Host "  $($PrecomputedNames.CSVName)" -ForegroundColor Yellow
			$CSVResponse = Read-Host "Enter name (Enter = accept, Ctrl+V = paste & edit)"
			
			if ([string]::IsNullOrWhiteSpace($CSVResponse)) {
				$ZipResult.CSV_Rename = $PrecomputedNames.CSVName
			}
			else {
				$ZipResult.CSV_Rename = $CSVResponse
			}

			Write-Log "  CSV will be renamed to: $($ZipResult.CSV_Rename)"

			# Rename the files
			$NewPDFName = "$($ZipResult.PDF_Rename).pdf"
			$NewCSVName = "$($ZipResult.CSV_Rename).txt"
			
			Rename-Item -Path $PdfFile.FullName -NewName $NewPDFName -Force
			Rename-Item -Path $CSVFile.FullName -NewName $NewCSVName -Force
			
			Write-Log "  Files renamed successfully"

			# Get new paths after renaming
			$NewPDFPath = Join-Path $ZipResult.ExtractPath $NewPDFName
			$NewCSVPath = Join-Path $ZipResult.ExtractPath $NewCSVName

			# Move files to final destinations
			Write-Log "  Moving files to final destinations..."
			$moveSuccess = Move-ProcessedFiles `
				-PDFPath $NewPDFPath `
				-CSVPath $NewCSVPath `
				-ACHFolder $ACHFolder `
				-InputTPAFolder $InputTPA

			if ($moveSuccess) {
				Write-Log "  Processing complete for this ZIP" -Level MATCH

                # [AMENDMENT 3] Clean up extract folder and source ZIP now that files are safely moved
                Write-Log "  Cleaning up temporary files..."
                $cleanedUp = Remove-ProcessedZipAndFolder `
                    -ZipPath $ZipResult.ZipPath `
                    -ExtractPath $ZipResult.ExtractPath

                if ($cleanedUp) {
                    Write-Log "  Downloads folder cleaned up successfully" -Level MATCH
                } else {
                    Write-Log "  WARNING: Some cleanup steps failed - check Downloads folder manually" -Level WARNING
                }
			}

		}
		catch {
			Write-Log "  Step 3 failed: $($_.Exception.Message)" -Level ERROR
			continue
		}
	}

    
} catch {
	Write-Log "Fatal error: $($_.Exception.Message)" -Level ERROR
	Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level ERROR
}

Write-Log ""
Write-Log "========================================"
Write-Log "Processing Complete"
Write-Log "========================================"
Write-Host ""
Write-Host "Log file saved to: $LogFile" -ForegroundColor Cyan
Write-Host ""
Read-Host "Press Enter to exit"
