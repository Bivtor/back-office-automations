<#
.SYNOPSIS
    TCG Remittance File Renamer + Fixed-Folder Processor
#>

param(
    [Parameter(Mandatory=$true)]
    [string]$ExcelFile,

    [Parameter(Mandatory=$true)]
    [string]$RemittanceFolder,

    [Parameter(Mandatory=$true)]
    [string]$Downloads,

    [string]$FileNameTemplate = "JNT [EMPLOYER_NAME] $(Get-Date -Format 'M-d') [AMOUNT]"
)

# ================= LOGGING =================

$LogFile = $null

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARNING','ERROR','MATCH','SUCCESS')]
        [string]$Level = 'INFO'
    )

    $Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $Line = "$Timestamp - $Level - $Message"

    switch ($Level) {
        'INFO'    { Write-Host $Line -ForegroundColor White }
        'WARNING' { Write-Host $Line -ForegroundColor Yellow }
        'ERROR'   { Write-Host $Line -ForegroundColor Red }
        'MATCH'   { Write-Host $Line -ForegroundColor Yellow }
        'SUCCESS' { Write-Host $Line -ForegroundColor Green }
    }

    Add-Content -Path $LogFile -Value $Line
}

# ================= EXCEL EXTRACTION (ALL DATA - single open) =================

function Extract-AllData {
    param([string]$ExcelPath)

    if (-not (Test-Path $ExcelPath)) {
        throw "Excel file not found: $ExcelPath"
    }

    $Excel    = $null
    $Workbook = $null

    try {
        Write-Log "Attempting to open Excel..."

        # Retry logic for Excel COM object
        $MaxRetries = 5
        $RetryCount = 0
        $Success    = $false

        while (-not $Success -and $RetryCount -lt $MaxRetries) {
            try {
                Write-Log "Creating Excel COM object (attempt $($RetryCount + 1) of $MaxRetries)..."
                $Excel = New-Object -ComObject Excel.Application
                Start-Sleep -Seconds 2

                $Excel.Visible           = $false
                $Excel.DisplayAlerts     = $false
                $Excel.AskToUpdateLinks  = $false
                $Excel.AutomationSecurity = 3
                $Excel.ScreenUpdating    = $false

                Write-Log "Opening workbook: $ExcelPath"
                Write-Log "Please wait, this may take a moment on slower systems..."

                $Workbook = $Excel.Workbooks.Open($ExcelPath, 0, $true)
                Start-Sleep -Seconds 3

                $null = $Workbook.Worksheets.Count  # throws if not ready

                $Success = $true
                Write-Log "Excel workbook opened successfully"
            }
            catch {
                $RetryCount++
                Write-Log "Attempt $RetryCount failed: $($_.Exception.Message)" -Level WARNING

                if ($RetryCount -lt $MaxRetries) {
                    Write-Log "Retrying in 5 seconds..." -Level WARNING
                    Start-Sleep -Seconds 5

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

                    $Excel    = $null
                    $Workbook = $null
                } else {
                    throw "Failed to open Excel after $MaxRetries attempts. Please close all Excel windows and try again."
                }
            }
        }

        $Worksheet = $Workbook.Worksheets.Item(1)
        $MaxRow    = $Worksheet.UsedRange.Rows.Count
        $TodayDate = Get-Date -Format "M-d"

        Write-Log "Reading data from worksheet ($MaxRow rows)..."

        $TCGEntries          = @()
        $FixedFolderEntries  = @()
        $OMNIEntries         = @()
        $SchoolsFirstEntries = @()
        $AFPEntries          = @()
        $AMEntries           = @()

        # Helper: read and clean an amount from col I or J of a given row
        $ReadAmount = {
            param($R)
            $Val = $Worksheet.Cells.Item($R, 9).Text
            if (-not $Val -or $Val -eq "") { $Val = $Worksheet.Cells.Item($R, 10).Text }
            $Val = $Val -replace '[$,]', ''
            if (-not $Val -or $Val -eq "") { return "0.00" }
            $Parsed = [decimal]0
            if ([decimal]::TryParse($Val, [ref]$Parsed)) {
                if ($Parsed % 1 -eq 0) { return $Parsed.ToString("0") }
                else                   { return $Parsed.ToString("0.00") }
            }
            return $Val
        }

        for ($Row = 1; $Row -le $MaxRow; $Row++) {
            if ($Row % 50 -eq 0) {
                Write-Log "Processing row $Row of $MaxRow..."
            }

            $Cell = $Worksheet.Cells.Item($Row, 2).Text

            # ---- TCG ----
            if ($Cell -and $Cell.StartsWith("TCG")) {
                $EmployerCell = $Worksheet.Cells.Item($Row + 2, 2).Text
                $EmployerName = ($EmployerCell -replace '^\d+', '').Trim()
                $Amount       = & $ReadAmount ($Row - 1)

                if ($EmployerCell -match '^(\d+)') {
                    $Digits = $matches[1]
                    if ($Digits.Length -gt 3) { $Digits = $Digits.Substring(3) }

                    $TCGEntries += [PSCustomObject]@{
                        AccountNumber = $Digits
                        EmployerName  = $EmployerName
                        Amount        = $Amount
                    }
                }
            }

            # ---- NBS ----
            if ($Cell -and $Cell -like "NATIONAL BENEFIT*") {
                $Amount = & $ReadAmount ($Row - 1)
                $FixedFolderEntries += [PSCustomObject]@{
                    RowNumber   = $Row
                    Company     = "NBS"
                    FolderPath  = "\\INTERNAL_FTP_SERVER\ftproot\nbs"
                    FileName    = "NBS $TodayDate $Amount"
                    Destination = $RemittanceFolder
                    Amount      = $Amount
                }
                Write-Log "Found NBS entry at row $Row with amount $Amount"
            }

            # ---- Mid Atlantic ----
            if ($Cell -and $Cell -like "*MID ATLANTIC*") {
                $Amount = & $ReadAmount ($Row - 1)
                $FixedFolderEntries += [PSCustomObject]@{
                    RowNumber   = $Row
                    Company     = "Mid Atlantic"
                    FolderPath  = "\\INTERNAL_FTP_SERVER\ftproot\adminpartners"
                    FileName    = "MA $TodayDate $Amount"
                    Destination = $RemittanceFolder
                    Amount      = $Amount
                }
                Write-Log "Found Mid Atlantic entry at row $Row with amount $Amount"
            }

            # ---- Penserv ----
            if ($Cell -and $Cell -like "*PENSERV*") {
                $Amount = & $ReadAmount ($Row - 1)
                $FixedFolderEntries += [PSCustomObject]@{
                    RowNumber   = $Row
                    Company     = "Penserv"
                    FolderPath  = "\\INTERNAL_FTP_SERVER\ftproot\penservfiles"
                    FileName    = "PS $TodayDate $Amount"
                    Destination = $RemittanceFolder
                    Amount      = $Amount
                }
                Write-Log "Found Penserv entry at row $Row with amount $Amount"
            }

            # ---- OMNI ----
            if ($Cell -and $Cell.StartsWith("OMNI FINANCIAL,")) {
                $Amount = & $ReadAmount ($Row - 1)
                $OMNIEntries += [PSCustomObject]@{
                    RowNumber   = $Row
                    Company     = "OMNI"
                    FileName    = "OMNI $TodayDate $Amount"
                    Destination = $RemittanceFolder
                    Amount      = $Amount
                }
                Write-Log "Found OMNI entry at row $Row with amount $Amount"
            }

            # ---- SchoolsFirst ----
            if ($Cell -and $Cell -like "SchoolsFirst FCU,*") {
                $Amount = & $ReadAmount ($Row - 1)
                $SchoolsFirstEntries += [PSCustomObject]@{
                    RowNumber   = $Row
                    Company     = "SF"
                    FileName    = "SF $TodayDate $Amount"
                    Destination = $RemittanceFolder
                    Amount      = $Amount
                }
                Write-Log "Found SchoolsFirst entry at row $Row with amount $Amount"
            }

            # ---- AFPlanServ ----
            if ($Cell -and $Cell -like "AFPLANSERV - AFA,*") {
                $Amount = & $ReadAmount ($Row - 1)
                $AFPEntries += [PSCustomObject]@{
                    RowNumber   = $Row
                    Company     = "AFP"
                    FileName    = "AFP $TodayDate $Amount"
                    Destination = $RemittanceFolder
                    Amount      = $Amount
                }
                Write-Log "Found AFPlanServ entry at row $Row with amount $Amount"
            }

            # ---- Alta Montclair ----
            if ($Cell -and $Cell -like "ALTA MONTCLAIR,*") {
                $Amount = & $ReadAmount ($Row - 1)
                $AMEntries += [PSCustomObject]@{
                    RowNumber   = $Row
                    Company     = "AM"
                    FileName    = "AM $TodayDate $Amount"
                    Destination = $RemittanceFolder
                    Amount      = $Amount
                }
                Write-Log "Found Alta Montclair entry at row $Row with amount $Amount"
            }
        }

        Write-Log "Extraction complete - TCG: $($TCGEntries.Count), Fixed-Folder: $($FixedFolderEntries.Count), OMNI: $($OMNIEntries.Count), SchoolsFirst: $($SchoolsFirstEntries.Count), AFPlanServ: $($AFPEntries.Count), Alta Montclair: $($AMEntries.Count)"

        return [PSCustomObject]@{
            TCG          = $TCGEntries
            FixedFolder  = $FixedFolderEntries
            OMNI         = $OMNIEntries
            SchoolsFirst = $SchoolsFirstEntries
            AFP          = $AFPEntries
            AM           = $AMEntries
        }
    }
    finally {
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

# ================= FILE HELPERS =================

function Find-MatchingFile {
    param(
        [string]$DownloadsFolder,
        [string]$AccountNumber
    )

    $Pattern = "ListBill_0${AccountNumber}_*.txt"
    $Files = Get-ChildItem -Path $DownloadsFolder -Filter $Pattern -ErrorAction SilentlyContinue

    if ($Files.Count -eq 0) {
        Write-Log "No file found for account $AccountNumber in Downloads" -Level WARNING
        return $null
    }

    return $Files[0].FullName
}

function Get-UniqueFileName {
    param(
        [string]$Directory,
        [string]$BaseFileName
    )

    $Candidate = Join-Path $Directory "$BaseFileName.txt"
    if (-not (Test-Path $Candidate)) {
        return "$BaseFileName.txt"
    }

    $Counter = 2
    while ($true) {
        if ($BaseFileName -match '^(JNT\s+\S+)(\s+.*)$') {
            $Name = "$($matches[1]) ($Counter)$($matches[2]).txt"
        } else {
            $Name = "$BaseFileName ($Counter).txt"
        }

        if (-not (Test-Path (Join-Path $Directory $Name))) {
            return $Name
        }

        $Counter++
        if ($Counter -gt 100) {
            return "$BaseFileName ($(Get-Date -Format 'HHmmss')).txt"
        }
    }
}

function Sanitize-FileName {
    param([string]$FileName)
    
    $Sanitized = $FileName -replace '[<>:"/\\|?*]', '_'
    $Sanitized = $Sanitized.Trim('. ')
    return $Sanitized
}

function Get-NewestFile {
    param([string]$FolderPath)
    
    if (-not (Test-Path $FolderPath)) {
        Write-Log "Folder not found: $FolderPath" -Level ERROR
        return $null
    }
    
    $Files = Get-ChildItem -Path $FolderPath -File | Sort-Object LastWriteTime -Descending
    
    if ($Files.Count -eq 0) {
        Write-Log "No files found in: $FolderPath" -Level WARNING
        return $null
    }
    
    return $Files[0]
}

function Sum-CSVColumn-NBS {
    param(
        [string]$FilePath,
        [int]$ColumnIndex
    )
    
    try {
        $Content  = Get-Content -Path $FilePath
        $Sum      = 0
        $RowCount = 0
        
        foreach ($Line in $Content) {
            $RowCount++
            $Columns = $Line -split ','
            
            if ($ColumnIndex -gt $Columns.Count) { continue }
            
            $Value = $Columns[$ColumnIndex - 1].Trim()
            $NumericValue = 0
            if ([decimal]::TryParse($Value, [ref]$NumericValue)) {
                $Sum += $NumericValue
            } else {
                Write-Log "DEBUG: Row $RowCount - Could not parse value: '$Value'" -Level WARNING
            }
        }
        
        return $Sum
    }
    catch {
        Write-Log "Error summing CSV column: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}

function Extract-AmountFromPipeDelimited-MidAtlantic {
    param([string]$FilePath)
    
    try {
        $Content = Get-Content -Path $FilePath -Raw
        $Parts   = $Content -split '\|'
        $Parts   = $Parts | Where-Object { $_.Trim() -ne '' }
        
        if ($Parts.Count -ge 2) {
            $Value = $Parts[$Parts.Count - 2].Trim()
            $NumericValue = 0
            if ([decimal]::TryParse($Value, [ref]$NumericValue)) {
                return $NumericValue
            } else {
                Write-Log "Could not parse Mid Atlantic value '$Value' as decimal" -Level WARNING
            }
        } else {
            Write-Log "Mid Atlantic file has only $($Parts.Count) non-empty parts, need at least 2" -Level WARNING
        }
        
        return $null
    }
    catch {
        Write-Log "Error extracting amount from Mid Atlantic file: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}

function Extract-AmountFromPipeDelimited-Penserv {
    param([string]$FilePath)
    
    try {
        $Content = Get-Content -Path $FilePath -Raw
        $Parts   = $Content -split '\|'
        
        if ($Parts.Count -ge 4) {
            $Value = $Parts[$Parts.Count - 4].Trim()
            
            if ([string]::IsNullOrWhiteSpace($Value)) {
                Write-Log "Found empty value at Penserv position (position 4 from end was empty)" -Level WARNING
                return $null
            }
            
            $NumericValue = 0
            if ([decimal]::TryParse($Value, [ref]$NumericValue)) {
                return $NumericValue
            } else {
                Write-Log "Could not parse Penserv value '$Value' as decimal" -Level WARNING
            }
        } else {
            Write-Log "Penserv file has only $($Parts.Count) pipe-delimited parts, need at least 4" -Level WARNING
        }
        
        return $null
    }
    catch {
        Write-Log "Error extracting amount from Penserv file: $($_.Exception.Message)" -Level ERROR
        return $null
    }
}

function Compare-Amounts {
    param(
        [decimal]$Amount1,
        [decimal]$Amount2,
        [decimal]$Tolerance = 0.01
    )
    
    return [Math]::Abs($Amount1 - $Amount2) -le $Tolerance
}

# ================= OMNI HELPERS =================

function Sum-OMNIColumns {
    param([string]$FilePath)

    $Excel    = $null
    $Workbook = $null

    try {
        $Excel = New-Object -ComObject Excel.Application
        Start-Sleep -Milliseconds 500

        $Excel.Visible           = $false
        $Excel.DisplayAlerts     = $false
        $Excel.AskToUpdateLinks  = $false
        $Excel.AutomationSecurity = 3
        $Excel.ScreenUpdating    = $false

        $Workbook  = $Excel.Workbooks.Open($FilePath, 0, $true)
        Start-Sleep -Seconds 1

        $Worksheet = $Workbook.Worksheets.Item(1)
        $UsedRange = $Worksheet.UsedRange
        $MaxRow    = $UsedRange.Rows.Count
        $MaxCol    = $UsedRange.Columns.Count

        $EEColIndex = $null
        $ERColIndex = $null
        $HeaderRow  = $null

        for ($R = 1; $R -le [Math]::Min($MaxRow, 10); $R++) {
            for ($C = 1; $C -le $MaxCol; $C++) {
                $HeaderText = $Worksheet.Cells.Item($R, $C).Text.Trim()
                if ($HeaderText -eq "NetFundsRemittedEE") {
                    $EEColIndex = $C
                    $HeaderRow  = $R
                }
                if ($HeaderText -eq "NetFundsRemittedER") {
                    $ERColIndex = $C
                    $HeaderRow  = $R
                }
            }
            if ($EEColIndex -and $ERColIndex) { break }
        }

        if ($null -eq $EEColIndex -or $null -eq $ERColIndex) {
            Write-Log "  Columns 'NetFundsRemittedEE' / 'NetFundsRemittedER' not found - skipping file" -Level WARNING
            return $null
        }

        Write-Log "  Found headers at row $HeaderRow (EE=col $EEColIndex, ER=col $ERColIndex)"

        $Total = [decimal]0

        for ($R = $HeaderRow + 1; $R -le $MaxRow; $R++) {
            $EEText = $Worksheet.Cells.Item($R, $EEColIndex).Text.Trim() -replace '[$,]', ''
            $ERText = $Worksheet.Cells.Item($R, $ERColIndex).Text.Trim() -replace '[$,]', ''

            $EEVal = [decimal]0
            $ERVal = [decimal]0

            [decimal]::TryParse($EEText, [ref]$EEVal) | Out-Null
            [decimal]::TryParse($ERText, [ref]$ERVal) | Out-Null

            $Total += $EEVal + $ERVal
        }

        Write-Log "  Summed EE + ER columns: $Total"
        return $Total
    }
    catch {
        Write-Log "  Error reading OMNI candidate file '$FilePath': $($_.Exception.Message)" -Level WARNING
        return $null
    }
    finally {
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
    }
}

# ================= SCHOOLSFIRST HELPERS =================

function Extract-AmountFromSchoolsFirst {
    <#
    .SYNOPSIS
        Opens an .xlsx file via Excel COM and reads the value in the last
        non-empty column of row 6. Returns $null on any failure.
    #>
    param([string]$FilePath)

    $Excel    = $null
    $Workbook = $null

    try {
        $Excel = New-Object -ComObject Excel.Application
        Start-Sleep -Milliseconds 500

        $Excel.Visible           = $false
        $Excel.DisplayAlerts     = $false
        $Excel.AskToUpdateLinks  = $false
        $Excel.AutomationSecurity = 3
        $Excel.ScreenUpdating    = $false

        $Workbook  = $Excel.Workbooks.Open($FilePath, 0, $true)
        Start-Sleep -Seconds 1

        $Worksheet = $Workbook.Worksheets.Item(1)
        $UsedRange = $Worksheet.UsedRange
        $MaxCol    = $UsedRange.Columns.Count

        # Walk row 6 from right to left to find last non-empty cell
        $TargetRow = 6
        $LastValue = $null

        for ($C = $MaxCol; $C -ge 1; $C--) {
            $CellText = $Worksheet.Cells.Item($TargetRow, $C).Text.Trim() -replace '[$,]', ''
            if (-not [string]::IsNullOrWhiteSpace($CellText)) {
                $NumericValue = [decimal]0
                if ([decimal]::TryParse($CellText, [ref]$NumericValue)) {
                    Write-Log "  SchoolsFirst: found value '$CellText' in row 6, col $C"
                    $LastValue = $NumericValue
                } else {
                    Write-Log "  SchoolsFirst: last non-empty cell in row 6 (col $C) contains non-numeric value '$CellText' - skipping" -Level WARNING
                }
                break   # Stop at the first non-empty cell from the right regardless
            }
        }

        return $LastValue
    }
    catch {
        Write-Log "  Error reading SchoolsFirst candidate file '$FilePath': $($_.Exception.Message)" -Level WARNING
        return $null
    }
    finally {
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
    }
}

# ================= AFPLANSERV HELPERS =================

function Extract-AmountFromAFPlanServ {
    <#
    .SYNOPSIS
        Opens an .xls file via Excel COM and reads the value from the last
        populated row in the merged column AE (col 31 / AE:AH merge).
        In Excel COM, merged cells store their value in the top-left cell of
        the merge (AE), so we walk column 31 from the bottom of the used range
        upward until we find a non-empty numeric value.
        Returns $null on any failure.
    #>
    param([string]$FilePath)

    $Excel    = $null
    $Workbook = $null

    try {
        $Excel = New-Object -ComObject Excel.Application
        Start-Sleep -Milliseconds 500

        $Excel.Visible           = $false
        $Excel.DisplayAlerts     = $false
        $Excel.AskToUpdateLinks  = $false
        $Excel.AutomationSecurity = 3
        $Excel.ScreenUpdating    = $false

        $Workbook  = $Excel.Workbooks.Open($FilePath, 0, $true)
        Start-Sleep -Seconds 1

        $Worksheet = $Workbook.Worksheets.Item(1)
        $UsedRange = $Worksheet.UsedRange
        $MaxRow    = $UsedRange.Rows.Count

        # Column AE = 31 (A=1 … Z=26, AA=27, AB=28, AC=29, AD=30, AE=31)
        $TargetCol = 31
        $LastValue = $null

        Write-Log "  AFPlanServ: scanning column AE (col 31) from row $MaxRow upward for last numeric value"

        for ($R = $MaxRow; $R -ge 1; $R--) {
            # Use the MergeArea's Value2 so we read the merge's stored value
            # even when $R points to a non-anchor row of a merged block.
            $CellObj  = $Worksheet.Cells.Item($R, $TargetCol)
            $RawValue = $null

            try {
                # MergeArea.Cells(1,1) is always the top-left anchor of the merge
                $RawValue = $CellObj.MergeArea.Cells.Item(1, 1).Value2
            } catch {
                # Fallback if MergeArea isn't available (non-merged cell)
                $RawValue = $CellObj.Value2
            }

            if ($null -eq $RawValue -or [string]::IsNullOrWhiteSpace([string]$RawValue)) {
                continue
            }

            $CellText = ([string]$RawValue).Trim() -replace '[$,]', ''

            $NumericValue = [decimal]0
            if ([decimal]::TryParse($CellText, [ref]$NumericValue)) {
                Write-Log "  AFPlanServ: found value '$CellText' in col AE, row $R"
                $LastValue = $NumericValue
                break
            } else {
                Write-Log "  AFPlanServ: non-numeric value '$CellText' in col AE row $R - continuing upward" -Level WARNING
            }
        }

        if ($null -eq $LastValue) {
            Write-Log "  AFPlanServ: no numeric value found in column AE" -Level WARNING
        }

        return $LastValue
    }
    catch {
        Write-Log "  Error reading AFPlanServ candidate file '$FilePath': $($_.Exception.Message)" -Level WARNING
        return $null
    }
    finally {
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
    }
}

# ================= ALTA MONTCLAIR HELPERS =================

function Extract-AmountFromAltaMontclair {
    <#
    .SYNOPSIS
        Opens an Alta Montclair .xls or .xlsx file via Excel COM and reads the
        value from the last (highest row number) numeric cell in column D.
        Walks column D from the bottom of the used range upward until it finds
        a non-empty numeric value. Returns $null on any failure.

        NOTE: Alta Montclair files often trigger Excel security warnings when
        opened manually. The COM settings below (AutomationSecurity = 3 =
        msoAutomationSecurityForceDisable, DisplayAlerts = false,
        AskToUpdateLinks = false) suppress those prompts for programmatic
        opens so this runs unattended.
    #>
    param([string]$FilePath)

    $Excel    = $null
    $Workbook = $null

    try {
        $Excel = New-Object -ComObject Excel.Application
        Start-Sleep -Milliseconds 500

        $Excel.Visible           = $false
        $Excel.DisplayAlerts     = $false
        $Excel.AskToUpdateLinks  = $false
        $Excel.AutomationSecurity = 3
        $Excel.ScreenUpdating    = $false

        $Workbook  = $Excel.Workbooks.Open($FilePath, 0, $true)
        Start-Sleep -Seconds 1

        $Worksheet = $Workbook.Worksheets.Item(1)
        $UsedRange = $Worksheet.UsedRange
        $MaxRow    = $UsedRange.Rows.Count

        # Column D = 4
        $TargetCol = 4
        $LastValue = $null

        Write-Log "  Alta Montclair: scanning column D (col 4) from row $MaxRow upward for last numeric value"

        for ($R = $MaxRow; $R -ge 1; $R--) {
            $CellText = $Worksheet.Cells.Item($R, $TargetCol).Text.Trim() -replace '[$,]', ''

            if ([string]::IsNullOrWhiteSpace($CellText)) {
                continue
            }

            $NumericValue = [decimal]0
            if ([decimal]::TryParse($CellText, [ref]$NumericValue)) {
                Write-Log "  Alta Montclair: found value '$CellText' in col D, row $R"
                $LastValue = $NumericValue
                break
            } else {
                Write-Log "  Alta Montclair: non-numeric value '$CellText' in col D row $R - continuing upward" -Level WARNING
            }
        }

        if ($null -eq $LastValue) {
            Write-Log "  Alta Montclair: no numeric value found in column D" -Level WARNING
        }

        return $LastValue
    }
    catch {
        Write-Log "  Error reading Alta Montclair candidate file '$FilePath': $($_.Exception.Message)" -Level WARNING
        return $null
    }
    finally {
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
    }
}


function Process-NBS {
    param(
        $Entry,
        [System.Collections.Generic.List[object]]$SummaryItems
    )
    
    Write-Log "========================================"
    Write-Log "Processing NBS entry..." -Level INFO
    Write-Log "Expected amount: $($Entry.Amount)"
    
    if (-not (Test-Path $Entry.FolderPath)) {
        Write-Log "NBS folder not found: $($Entry.FolderPath)" -Level ERROR
        return
    }
    
    $Files = Get-ChildItem -Path $Entry.FolderPath -File | Sort-Object LastWriteTime -Descending | Select-Object -First 5
    
    if ($Files.Count -eq 0) {
        Write-Log "No files found in NBS folder" -Level WARNING
        return
    }
    
    Write-Log "Found $($Files.Count) file(s) to check"
    
    $ExpectedAmount = [decimal]$Entry.Amount
    $MatchedFile    = $null
    
    for ($i = 0; $i -lt $Files.Count; $i++) {
        $CurrentFile = $Files[$i]
        Write-Log "Checking file $($i + 1) of $($Files.Count): $($CurrentFile.Name)"
        
        $Sum = Sum-CSVColumn-NBS -FilePath $CurrentFile.FullName -ColumnIndex 8
        
        if ($null -eq $Sum) {
            Write-Log "Failed to calculate sum from file - skipping" -Level WARNING
            continue
        }
        
        Write-Log "Calculated sum from column 8: $Sum"
        
        if (Compare-Amounts -Amount1 $Sum -Amount2 $ExpectedAmount) {
            Write-Log "Amount MATCHES! Using this file." -Level SUCCESS
            $MatchedFile = $CurrentFile
            break
        } else {
            Write-Log "Amount mismatch - Expected: $ExpectedAmount, Got: $Sum" -Level WARNING
        }
    }
    
    if ($null -eq $MatchedFile) {
        Write-Log "No matching file found after checking $($Files.Count) file(s)" -Level ERROR
        Write-Log "File NOT processed - no amount match found" -Level WARNING
        return
    }
    
    Write-Log "Proceeding with copy and move..." -Level INFO
    
    $DestPath = Join-Path $Entry.Destination "$($Entry.FileName).csv"
    $Counter  = 2
    while (Test-Path $DestPath) {
        $DestPath = Join-Path $Entry.Destination "$($Entry.FileName) ($Counter).csv"
        $Counter++
    }
    
    try {
        Copy-Item -Path $MatchedFile.FullName -Destination $DestPath -ErrorAction Stop
        Write-Log "File copied to: $DestPath" -Level SUCCESS
        
        $CompleteFolder = Join-Path $Entry.FolderPath "Complete"
        if (Test-Path $CompleteFolder) {
            $MoveDest = Join-Path $CompleteFolder $MatchedFile.Name
            $Counter  = 2
            while (Test-Path $MoveDest) {
                $BaseName  = [System.IO.Path]::GetFileNameWithoutExtension($MatchedFile.Name)
                $Extension = [System.IO.Path]::GetExtension($MatchedFile.Name)
                $MoveDest  = Join-Path $CompleteFolder "$BaseName ($Counter)$Extension"
                $Counter++
            }
            
            Move-Item -Path $MatchedFile.FullName -Destination $MoveDest -ErrorAction Stop
            Write-Log "Original file moved to: $MoveDest" -Level SUCCESS
        } else {
            Write-Log "Complete folder not found: $CompleteFolder - file not moved" -Level WARNING
        }

        $SummaryItems.Add([PSCustomObject]@{
            Company      = "NBS"
            DisplayName  = "NBS"
            FileName     = (Split-Path $DestPath -Leaf)
            Amount       = $ExpectedAmount
            AmountStr    = $Entry.Amount
        })
    }
    catch {
        Write-Log "Error during file operations: $($_.Exception.Message)" -Level ERROR
    }
}

function Process-MidAtlantic {
    param(
        $Entry,
        [System.Collections.Generic.List[object]]$SummaryItems
    )
    
    Write-Log "========================================"
    Write-Log "Processing Mid Atlantic entry..." -Level INFO
    Write-Log "Expected amount: $($Entry.Amount)"
    
    if (-not (Test-Path $Entry.FolderPath)) {
        Write-Log "Mid Atlantic folder not found: $($Entry.FolderPath)" -Level ERROR
        return
    }
    
    $Files = Get-ChildItem -Path $Entry.FolderPath -File | Sort-Object LastWriteTime -Descending | Select-Object -First 5
    
    if ($Files.Count -eq 0) {
        Write-Log "No files found in Mid Atlantic folder" -Level WARNING
        return
    }
    
    Write-Log "Found $($Files.Count) file(s) to check"
    
    $ExpectedAmount = [decimal]$Entry.Amount
    $MatchedFile    = $null
    
    for ($i = 0; $i -lt $Files.Count; $i++) {
        $CurrentFile = $Files[$i]
        Write-Log "Checking file $($i + 1) of $($Files.Count): $($CurrentFile.Name)"
        
        $ExtractedAmount = Extract-AmountFromPipeDelimited-MidAtlantic -FilePath $CurrentFile.FullName
        
        if ($null -eq $ExtractedAmount) {
            Write-Log "Failed to extract amount from file - skipping" -Level WARNING
            continue
        }
        
        Write-Log "Extracted amount from file: $ExtractedAmount"
        
        if (Compare-Amounts -Amount1 $ExtractedAmount -Amount2 $ExpectedAmount) {
            Write-Log "Amount MATCHES! Using this file." -Level SUCCESS
            $MatchedFile = $CurrentFile
            break
        } else {
            Write-Log "Amount mismatch - Expected: $ExpectedAmount, Got: $ExtractedAmount" -Level WARNING
        }
    }
    
    if ($null -eq $MatchedFile) {
        Write-Log "No matching file found after checking $($Files.Count) file(s)" -Level ERROR
        Write-Log "File NOT processed - no amount match found" -Level WARNING
        return
    }
    
    Write-Log "Proceeding with copy..." -Level INFO
    
    $DestPath = Join-Path $Entry.Destination "$($Entry.FileName).txt"
    $Counter  = 2
    while (Test-Path $DestPath) {
        $DestPath = Join-Path $Entry.Destination "$($Entry.FileName) ($Counter).txt"
        $Counter++
    }
    
    try {
        Copy-Item -Path $MatchedFile.FullName -Destination $DestPath -ErrorAction Stop
        Write-Log "File copied to: $DestPath" -Level SUCCESS

        $SummaryItems.Add([PSCustomObject]@{
            Company     = "Mid Atlantic"
            DisplayName = "Mid Atlantic"
            FileName    = (Split-Path $DestPath -Leaf)
            Amount      = $ExpectedAmount
            AmountStr   = $Entry.Amount
        })
    }
    catch {
        Write-Log "Error during file copy: $($_.Exception.Message)" -Level ERROR
    }
}

function Process-Penserv {
    param(
        $Entry,
        [System.Collections.Generic.List[object]]$SummaryItems
    )
    
    Write-Log "========================================"
    Write-Log "Processing Penserv entry..." -Level INFO
    Write-Log "Expected amount: $($Entry.Amount)"
    
    if (-not (Test-Path $Entry.FolderPath)) {
        Write-Log "Penserv folder not found: $($Entry.FolderPath)" -Level ERROR
        return
    }
    
    $Files = Get-ChildItem -Path $Entry.FolderPath -File | Sort-Object LastWriteTime -Descending | Select-Object -First 5
    
    if ($Files.Count -eq 0) {
        Write-Log "No files found in Penserv folder" -Level WARNING
        return
    }
    
    Write-Log "Found $($Files.Count) file(s) to check"
    
    $ExpectedAmount = [decimal]$Entry.Amount
    $MatchedFile    = $null
    
    for ($i = 0; $i -lt $Files.Count; $i++) {
        $CurrentFile = $Files[$i]
        Write-Log "Checking file $($i + 1) of $($Files.Count): $($CurrentFile.Name)"
        
        $ExtractedAmount = Extract-AmountFromPipeDelimited-Penserv -FilePath $CurrentFile.FullName
        
        if ($null -eq $ExtractedAmount) {
            Write-Log "Failed to extract amount from file - skipping" -Level WARNING
            continue
        }
        
        Write-Log "Extracted amount from file: $ExtractedAmount"
        
        if (Compare-Amounts -Amount1 $ExtractedAmount -Amount2 $ExpectedAmount) {
            Write-Log "Amount MATCHES! Using this file." -Level SUCCESS
            $MatchedFile = $CurrentFile
            break
        } else {
            Write-Log "Amount mismatch - Expected: $ExpectedAmount, Got: $ExtractedAmount" -Level WARNING
        }
    }
    
    if ($null -eq $MatchedFile) {
        Write-Log "No matching file found after checking $($Files.Count) file(s)" -Level ERROR
        Write-Log "File NOT processed - no amount match found" -Level WARNING
        return
    }
    
    Write-Log "Proceeding with copy..." -Level INFO
    
    $DestPath = Join-Path $Entry.Destination "$($Entry.FileName).txt"
    $Counter  = 2
    while (Test-Path $DestPath) {
        $DestPath = Join-Path $Entry.Destination "$($Entry.FileName) ($Counter).txt"
        $Counter++
    }
    
    try {
        Copy-Item -Path $MatchedFile.FullName -Destination $DestPath -ErrorAction Stop
        Write-Log "File copied to: $DestPath" -Level SUCCESS

        $SummaryItems.Add([PSCustomObject]@{
            Company     = "Penserv"
            DisplayName = "Penserv"
            FileName    = (Split-Path $DestPath -Leaf)
            Amount      = $ExpectedAmount
            AmountStr   = $Entry.Amount
        })
    }
    catch {
        Write-Log "Error during file copy: $($_.Exception.Message)" -Level ERROR
    }
}

function Process-OMNI {
    param(
        [array]$Entries,
        [string]$DownloadsFolder,
        [System.Collections.Generic.List[object]]$SummaryItems
    )

    Write-Log ""
    Write-Log "========================================"
    Write-Log "OMNI PROCESSING STARTED"
    Write-Log "========================================"
    Write-Log ""

    if ($Entries.Count -eq 0) {
        Write-Log "No OMNI entries found in Excel file" -Level WARNING
        return
    }

    Write-Log "Found $($Entries.Count) OMNI entries to match"

    $Today    = (Get-Date).Date
    $XlsFiles = Get-ChildItem -Path $DownloadsFolder -File |
                Where-Object { $_.Extension -in @('.xls', '.xlsx') -and $_.LastWriteTime.Date -eq $Today } |
                Sort-Object LastWriteTime -Descending

    if ($XlsFiles.Count -eq 0) {
        Write-Log "No .xls/.xlsx files downloaded today found in Downloads folder - OMNI files may not have been downloaded yet" -Level ERROR
        return
    }

    Write-Log "Found $($XlsFiles.Count) spreadsheet file(s) downloaded today to check (newest first)"

    $UnmatchedEntries = [System.Collections.Generic.List[object]]::new()
    foreach ($Entry in $Entries) {
        $UnmatchedEntries.Add([PSCustomObject]@{
            Amount      = [decimal]$Entry.Amount
            AmountStr   = $Entry.Amount
            FileName    = $Entry.FileName
            Destination = $Entry.Destination
            Matched     = $false
        })
    }

    $RemainingUnmatched = $Entries.Count

    foreach ($XlsFile in $XlsFiles) {

        # Stop early once every entry has been matched
        if ($RemainingUnmatched -le 0) {
            Write-Log "All OMNI entries matched - stopping file scan early." -Level SUCCESS
            break
        }

        Write-Log "----------------------------------------"
        Write-Log "Checking file: $($XlsFile.Name)"

        $FileSum = Sum-OMNIColumns -FilePath $XlsFile.FullName

        if ($null -eq $FileSum) {
            Write-Log "  Could not read OMNI columns - skipping" -Level WARNING
            continue
        }

        Write-Log "  Sum of EE + ER columns: $FileSum"

        $MatchedEntry = $null
        foreach ($Entry in $UnmatchedEntries) {
            if (-not $Entry.Matched -and (Compare-Amounts -Amount1 $FileSum -Amount2 $Entry.Amount)) {
                $MatchedEntry = $Entry
                break
            }
        }

        if ($null -eq $MatchedEntry) {
            Write-Log "  Sum $FileSum does not match any remaining unmatched OMNI entry - skipping file" -Level WARNING
            continue
        }

        Write-Log "  Matched entry with expected amount $($MatchedEntry.Amount)!" -Level SUCCESS
        $MatchedEntry.Matched = $true
        $RemainingUnmatched--

        $OriginalExtension = $XlsFile.Extension
        $NewFileName       = "$($MatchedEntry.FileName)$OriginalExtension"
        $NewNameSanitized  = Sanitize-FileName -FileName $NewFileName

        $RenamedPath = Join-Path $DownloadsFolder $NewNameSanitized
        $Counter = 2
        while (Test-Path $RenamedPath) {
            $Base        = [System.IO.Path]::GetFileNameWithoutExtension($NewNameSanitized)
            $Ext         = [System.IO.Path]::GetExtension($NewNameSanitized)
            $RenamedPath = Join-Path $DownloadsFolder "$Base ($Counter)$Ext"
            $Counter++
        }

        $DestPath = Join-Path $MatchedEntry.Destination (Split-Path $RenamedPath -Leaf)
        $Counter  = 2
        while (Test-Path $DestPath) {
            $Base     = [System.IO.Path]::GetFileNameWithoutExtension($NewNameSanitized)
            $Ext      = [System.IO.Path]::GetExtension($NewNameSanitized)
            $DestPath = Join-Path $MatchedEntry.Destination "$Base ($Counter)$Ext"
            $Counter++
        }

        try {
            Rename-Item -Path $XlsFile.FullName -NewName (Split-Path $RenamedPath -Leaf) -ErrorAction Stop
            Write-Log "  Renamed: $($XlsFile.Name) -> $(Split-Path $RenamedPath -Leaf)" -Level INFO

            Move-Item -Path $RenamedPath -Destination $DestPath -ErrorAction Stop
            Write-Log "  Moved to Input_TPA: $(Split-Path $DestPath -Leaf)" -Level SUCCESS

            $SummaryItems.Add([PSCustomObject]@{
                Company     = "OMNI"
                DisplayName = "OMNI"
                FileName    = (Split-Path $DestPath -Leaf)
                Amount      = $MatchedEntry.Amount
                AmountStr   = $MatchedEntry.AmountStr
            })
        }
        catch {
            Write-Log "  Error processing OMNI file: $($_.Exception.Message)" -Level ERROR
            $MatchedEntry.Matched = $false
            $RemainingUnmatched++
        }
    }

    Write-Log ""
    $Unresolved = $UnmatchedEntries | Where-Object { -not $_.Matched }
    $ResolvedCount = $UnmatchedEntries.Count - $Unresolved.Count
    if ($Unresolved.Count -gt 0) {
        Write-Log "The following OMNI entries were NOT matched to any downloaded file:" -Level ERROR
        foreach ($U in $Unresolved) {
            Write-Log "  - Expected amount: $($U.Amount)  -> target filename: $($U.FileName)" -Level ERROR
        }
        if ($ResolvedCount -gt 0) {
            Write-Log "$ResolvedCount of $($UnmatchedEntries.Count) OMNI entries were matched successfully." -Level WARNING
        }
    } elseif ($ResolvedCount -eq 0) {
        Write-Log "No OMNI entries were matched - no files moved." -Level ERROR
    } else {
        Write-Log "All OMNI entries matched and processed successfully." -Level SUCCESS
    }

    Write-Log ""
    Write-Log "========================================"
    Write-Log "OMNI PROCESSING COMPLETE"
    Write-Log "========================================"
}

function Process-SchoolsFirst {
    param(
        [array]$Entries,
        [string]$DownloadsFolder,
        [System.Collections.Generic.List[object]]$SummaryItems
    )

    Write-Log ""
    Write-Log "========================================"
    Write-Log "SCHOOLSFIRST PROCESSING STARTED"
    Write-Log "========================================"
    Write-Log ""

    if ($Entries.Count -eq 0) {
        Write-Log "No SchoolsFirst entries found in Excel file" -Level WARNING
        return
    }

    Write-Log "Found $($Entries.Count) SchoolsFirst entries to match"

    # Collect ALL .xlsx files in Downloads, newest first
    # (SchoolsFirst files come from email - no date filter, user may have downloaded earlier)
    $XlsxFiles = Get-ChildItem -Path $DownloadsFolder -File |
                 Where-Object { $_.Extension -eq '.xlsx' } |
                 Sort-Object LastWriteTime -Descending

    if ($XlsxFiles.Count -eq 0) {
        Write-Log "No .xlsx files found in Downloads folder" -Level ERROR
        return
    }

    Write-Log "Found $($XlsxFiles.Count) .xlsx file(s) to check (newest first)"

    $UnmatchedEntries = [System.Collections.Generic.List[object]]::new()
    foreach ($Entry in $Entries) {
        $UnmatchedEntries.Add([PSCustomObject]@{
            Amount      = [decimal]$Entry.Amount
            AmountStr   = $Entry.Amount
            FileName    = $Entry.FileName
            Destination = $Entry.Destination
            Matched     = $false
        })
    }

    $RemainingUnmatched = $Entries.Count

    foreach ($XlsxFile in $XlsxFiles) {

        # Stop early once every entry has been matched
        if ($RemainingUnmatched -le 0) {
            Write-Log "All SchoolsFirst entries matched - stopping file scan early." -Level SUCCESS
            break
        }

        Write-Log "----------------------------------------"
        Write-Log "Checking file: $($XlsxFile.Name)"

        $FileAmount = Extract-AmountFromSchoolsFirst -FilePath $XlsxFile.FullName

        if ($null -eq $FileAmount) {
            Write-Log "  Could not read amount from row 6 - skipping" -Level WARNING
            continue
        }

        Write-Log "  Amount from row 6 last column: $FileAmount"

        $MatchedEntry = $null
        foreach ($Entry in $UnmatchedEntries) {
            if (-not $Entry.Matched -and (Compare-Amounts -Amount1 $FileAmount -Amount2 $Entry.Amount)) {
                $MatchedEntry = $Entry
                break
            }
        }

        if ($null -eq $MatchedEntry) {
            Write-Log "  Amount $FileAmount does not match any remaining unmatched SchoolsFirst entry - skipping file" -Level WARNING
            continue
        }

        Write-Log "  Matched entry with expected amount $($MatchedEntry.Amount)!" -Level SUCCESS
        $MatchedEntry.Matched = $true
        $RemainingUnmatched--

        $OriginalExtension = $XlsxFile.Extension
        $NewFileName       = "$($MatchedEntry.FileName)$OriginalExtension"
        $NewNameSanitized  = Sanitize-FileName -FileName $NewFileName

        $RenamedPath = Join-Path $DownloadsFolder $NewNameSanitized
        $Counter = 2
        while (Test-Path $RenamedPath) {
            $Base        = [System.IO.Path]::GetFileNameWithoutExtension($NewNameSanitized)
            $Ext         = [System.IO.Path]::GetExtension($NewNameSanitized)
            $RenamedPath = Join-Path $DownloadsFolder "$Base ($Counter)$Ext"
            $Counter++
        }

        $DestPath = Join-Path $MatchedEntry.Destination (Split-Path $RenamedPath -Leaf)
        $Counter  = 2
        while (Test-Path $DestPath) {
            $Base     = [System.IO.Path]::GetFileNameWithoutExtension($NewNameSanitized)
            $Ext      = [System.IO.Path]::GetExtension($NewNameSanitized)
            $DestPath = Join-Path $MatchedEntry.Destination "$Base ($Counter)$Ext"
            $Counter++
        }

        try {
            Rename-Item -Path $XlsxFile.FullName -NewName (Split-Path $RenamedPath -Leaf) -ErrorAction Stop
            Write-Log "  Renamed: $($XlsxFile.Name) -> $(Split-Path $RenamedPath -Leaf)" -Level INFO

            Move-Item -Path $RenamedPath -Destination $DestPath -ErrorAction Stop
            Write-Log "  Moved to Input_TPA: $(Split-Path $DestPath -Leaf)" -Level SUCCESS

            $SummaryItems.Add([PSCustomObject]@{
                Company     = "SF"
                DisplayName = "SchoolsFirst"
                FileName    = (Split-Path $DestPath -Leaf)
                Amount      = $MatchedEntry.Amount
                AmountStr   = $MatchedEntry.AmountStr
            })
        }
        catch {
            Write-Log "  Error processing SchoolsFirst file: $($_.Exception.Message)" -Level ERROR
            $MatchedEntry.Matched = $false
            $RemainingUnmatched++
        }
    }

    Write-Log ""
    $Unresolved = $UnmatchedEntries | Where-Object { -not $_.Matched }
    $ResolvedCount = $UnmatchedEntries.Count - $Unresolved.Count
    if ($Unresolved.Count -gt 0) {
        Write-Log "The following SchoolsFirst entries were NOT matched to any downloaded file:" -Level ERROR
        foreach ($U in $Unresolved) {
            Write-Log "  - Expected amount: $($U.Amount)  -> target filename: $($U.FileName)" -Level ERROR
        }
        if ($ResolvedCount -gt 0) {
            Write-Log "$ResolvedCount of $($UnmatchedEntries.Count) SchoolsFirst entries were matched successfully." -Level WARNING
        }
    } elseif ($ResolvedCount -eq 0) {
        Write-Log "No SchoolsFirst entries were matched - no files moved." -Level ERROR
    } else {
        Write-Log "All SchoolsFirst entries matched and processed successfully." -Level SUCCESS
    }

    Write-Log ""
    Write-Log "========================================"
    Write-Log "SCHOOLSFIRST PROCESSING COMPLETE"
    Write-Log "========================================"
}

function Process-AFPlanServ {
    param(
        [array]$Entries,
        [string]$DownloadsFolder,
        [System.Collections.Generic.List[object]]$SummaryItems
    )

    Write-Log ""
    Write-Log "========================================"
    Write-Log "AFPLANSERV PROCESSING STARTED"
    Write-Log "========================================"
    Write-Log ""

    if ($Entries.Count -eq 0) {
        Write-Log "No AFPlanServ entries found in Excel file" -Level WARNING
        return
    }

    Write-Log "Found $($Entries.Count) AFPlanServ entries to match"

    # Check ALL .xls files in Downloads, newest first
    $XlsFiles = Get-ChildItem -Path $DownloadsFolder -File |
                Where-Object { $_.Extension -eq '.xls' } |
                Sort-Object LastWriteTime -Descending

    if ($XlsFiles.Count -eq 0) {
        Write-Log "No .xls files found in Downloads folder" -Level ERROR
        return
    }

    Write-Log "Found $($XlsFiles.Count) .xls file(s) to check (newest first)"

    $UnmatchedEntries = [System.Collections.Generic.List[object]]::new()
    foreach ($Entry in $Entries) {
        $UnmatchedEntries.Add([PSCustomObject]@{
            Amount      = [decimal]$Entry.Amount
            AmountStr   = $Entry.Amount
            FileName    = $Entry.FileName
            Destination = $Entry.Destination
            Matched     = $false
        })
    }

    $RemainingUnmatched = $Entries.Count

    foreach ($XlsFile in $XlsFiles) {

        if ($RemainingUnmatched -le 0) {
            Write-Log "All AFPlanServ entries matched - stopping file scan early." -Level SUCCESS
            break
        }

        Write-Log "----------------------------------------"
        Write-Log "Checking file: $($XlsFile.Name)"

        $FileAmount = Extract-AmountFromAFPlanServ -FilePath $XlsFile.FullName

        if ($null -eq $FileAmount) {
            Write-Log "  Could not read amount from column AE - skipping" -Level WARNING
            continue
        }

        Write-Log "  Amount from column AE (last populated row): $FileAmount"

        $MatchedEntry = $null
        foreach ($Entry in $UnmatchedEntries) {
            if (-not $Entry.Matched -and (Compare-Amounts -Amount1 $FileAmount -Amount2 $Entry.Amount)) {
                $MatchedEntry = $Entry
                break
            }
        }

        if ($null -eq $MatchedEntry) {
            Write-Log "  Amount $FileAmount does not match any remaining unmatched AFPlanServ entry - skipping file" -Level WARNING
            continue
        }

        Write-Log "  Matched entry with expected amount $($MatchedEntry.Amount)!" -Level SUCCESS
        $MatchedEntry.Matched = $true
        $RemainingUnmatched--

        $OriginalExtension = $XlsFile.Extension
        $NewFileName       = "$($MatchedEntry.FileName)$OriginalExtension"
        $NewNameSanitized  = Sanitize-FileName -FileName $NewFileName

        $RenamedPath = Join-Path $DownloadsFolder $NewNameSanitized
        $Counter = 2
        while (Test-Path $RenamedPath) {
            $Base        = [System.IO.Path]::GetFileNameWithoutExtension($NewNameSanitized)
            $Ext         = [System.IO.Path]::GetExtension($NewNameSanitized)
            $RenamedPath = Join-Path $DownloadsFolder "$Base ($Counter)$Ext"
            $Counter++
        }

        $DestPath = Join-Path $MatchedEntry.Destination (Split-Path $RenamedPath -Leaf)
        $Counter  = 2
        while (Test-Path $DestPath) {
            $Base     = [System.IO.Path]::GetFileNameWithoutExtension($NewNameSanitized)
            $Ext      = [System.IO.Path]::GetExtension($NewNameSanitized)
            $DestPath = Join-Path $MatchedEntry.Destination "$Base ($Counter)$Ext"
            $Counter++
        }

        try {
            Rename-Item -Path $XlsFile.FullName -NewName (Split-Path $RenamedPath -Leaf) -ErrorAction Stop
            Write-Log "  Renamed: $($XlsFile.Name) -> $(Split-Path $RenamedPath -Leaf)" -Level INFO

            Move-Item -Path $RenamedPath -Destination $DestPath -ErrorAction Stop
            Write-Log "  Moved to Input_TPA: $(Split-Path $DestPath -Leaf)" -Level SUCCESS

            $SummaryItems.Add([PSCustomObject]@{
                Company     = "AFP"
                DisplayName = "AFPlanServ"
                FileName    = (Split-Path $DestPath -Leaf)
                Amount      = $MatchedEntry.Amount
                AmountStr   = $MatchedEntry.AmountStr
            })
        }
        catch {
            Write-Log "  Error processing AFPlanServ file: $($_.Exception.Message)" -Level ERROR
            $MatchedEntry.Matched = $false
            $RemainingUnmatched++
        }
    }

    Write-Log ""
    $Unresolved = $UnmatchedEntries | Where-Object { -not $_.Matched }
    if ($Unresolved.Count -gt 0) {
        Write-Log "The following AFPlanServ entries were NOT matched to any downloaded file:" -Level ERROR
        foreach ($U in $Unresolved) {
            Write-Log "  - Expected amount: $($U.Amount)  -> target filename: $($U.FileName)" -Level ERROR
        }
    } else {
        Write-Log "All AFPlanServ entries matched and processed successfully." -Level SUCCESS
    }

    Write-Log ""
    Write-Log "========================================"
    Write-Log "AFPLANSERV PROCESSING COMPLETE"
    Write-Log "========================================"
}

function Process-AltaMontclair {
    param(
        [array]$Entries,
        [string]$DownloadsFolder,
        [System.Collections.Generic.List[object]]$SummaryItems
    )

    Write-Log ""
    Write-Log "========================================"
    Write-Log "ALTA MONTCLAIR PROCESSING STARTED"
    Write-Log "========================================"
    Write-Log ""

    if ($Entries.Count -eq 0) {
        Write-Log "No Alta Montclair entries found in Excel file" -Level WARNING
        return
    }

    Write-Log "Found $($Entries.Count) Alta Montclair entries to match"

    # Collect ALL .xls and .xlsx files in Downloads, newest first
    # (Alta Montclair files come from email - no date filter, user may have downloaded earlier)
    $XlsFiles = Get-ChildItem -Path $DownloadsFolder -File |
                Where-Object { $_.Extension -in @('.xls', '.xlsx') } |
                Sort-Object LastWriteTime -Descending

    if ($XlsFiles.Count -eq 0) {
        Write-Log "No .xls/.xlsx files found in Downloads folder" -Level ERROR
        return
    }

    Write-Log "Found $($XlsFiles.Count) spreadsheet file(s) to check (newest first)"

    $UnmatchedEntries = [System.Collections.Generic.List[object]]::new()
    foreach ($Entry in $Entries) {
        $UnmatchedEntries.Add([PSCustomObject]@{
            Amount      = [decimal]$Entry.Amount
            AmountStr   = $Entry.Amount
            FileName    = $Entry.FileName
            Destination = $Entry.Destination
            Matched     = $false
        })
    }

    $RemainingUnmatched = $Entries.Count

    foreach ($XlsFile in $XlsFiles) {

        if ($RemainingUnmatched -le 0) {
            Write-Log "All Alta Montclair entries matched - stopping file scan early." -Level SUCCESS
            break
        }

        Write-Log "----------------------------------------"
        Write-Log "Checking file: $($XlsFile.Name)"

        $FileAmount = Extract-AmountFromAltaMontclair -FilePath $XlsFile.FullName

        if ($null -eq $FileAmount) {
            Write-Log "  Could not read amount from column D - skipping" -Level WARNING
            continue
        }

        Write-Log "  Amount from column D (last populated row): $FileAmount"

        $MatchedEntry = $null
        foreach ($Entry in $UnmatchedEntries) {
            if (-not $Entry.Matched -and (Compare-Amounts -Amount1 $FileAmount -Amount2 $Entry.Amount)) {
                $MatchedEntry = $Entry
                break
            }
        }

        if ($null -eq $MatchedEntry) {
            Write-Log "  Amount $FileAmount does not match any remaining unmatched Alta Montclair entry - skipping file" -Level WARNING
            continue
        }

        Write-Log "  Matched entry with expected amount $($MatchedEntry.Amount)!" -Level SUCCESS
        $MatchedEntry.Matched = $true
        $RemainingUnmatched--

        $OriginalExtension = $XlsFile.Extension
        $NewFileName       = "$($MatchedEntry.FileName)$OriginalExtension"
        $NewNameSanitized  = Sanitize-FileName -FileName $NewFileName

        $RenamedPath = Join-Path $DownloadsFolder $NewNameSanitized
        $Counter = 2
        while (Test-Path $RenamedPath) {
            $Base        = [System.IO.Path]::GetFileNameWithoutExtension($NewNameSanitized)
            $Ext         = [System.IO.Path]::GetExtension($NewNameSanitized)
            $RenamedPath = Join-Path $DownloadsFolder "$Base ($Counter)$Ext"
            $Counter++
        }

        $DestPath = Join-Path $MatchedEntry.Destination (Split-Path $RenamedPath -Leaf)
        $Counter  = 2
        while (Test-Path $DestPath) {
            $Base     = [System.IO.Path]::GetFileNameWithoutExtension($NewNameSanitized)
            $Ext      = [System.IO.Path]::GetExtension($NewNameSanitized)
            $DestPath = Join-Path $MatchedEntry.Destination "$Base ($Counter)$Ext"
            $Counter++
        }

        try {
            Rename-Item -Path $XlsFile.FullName -NewName (Split-Path $RenamedPath -Leaf) -ErrorAction Stop
            Write-Log "  Renamed: $($XlsFile.Name) -> $(Split-Path $RenamedPath -Leaf)" -Level INFO

            Move-Item -Path $RenamedPath -Destination $DestPath -ErrorAction Stop
            Write-Log "  Moved to Input_TPA: $(Split-Path $DestPath -Leaf)" -Level SUCCESS

            $SummaryItems.Add([PSCustomObject]@{
                Company     = "AM"
                DisplayName = "Alta Montclair"
                FileName    = (Split-Path $DestPath -Leaf)
                Amount      = $MatchedEntry.Amount
                AmountStr   = $MatchedEntry.AmountStr
            })
        }
        catch {
            Write-Log "  Error processing Alta Montclair file: $($_.Exception.Message)" -Level ERROR
            $MatchedEntry.Matched = $false
            $RemainingUnmatched++
        }
    }

    Write-Log ""
    $Unresolved = $UnmatchedEntries | Where-Object { -not $_.Matched }
    $ResolvedCount = $UnmatchedEntries.Count - $Unresolved.Count
    if ($Unresolved.Count -gt 0) {
        Write-Log "The following Alta Montclair entries were NOT matched to any downloaded file:" -Level ERROR
        foreach ($U in $Unresolved) {
            Write-Log "  - Expected amount: $($U.Amount)  -> target filename: $($U.FileName)" -Level ERROR
        }
        if ($ResolvedCount -gt 0) {
            Write-Log "$ResolvedCount of $($UnmatchedEntries.Count) Alta Montclair entries were matched successfully." -Level WARNING
        }
    } elseif ($ResolvedCount -eq 0) {
        Write-Log "No Alta Montclair entries were matched - no files moved." -Level ERROR
    } else {
        Write-Log "All Alta Montclair entries matched and processed successfully." -Level SUCCESS
    }

    Write-Log ""
    Write-Log "========================================"
    Write-Log "ALTA MONTCLAIR PROCESSING COMPLETE"
    Write-Log "========================================"
}



function Rename-Files {
    param(
        [array]$Entries,
        [string]$DownloadsFolder,
        [string]$RemittanceFolder,
        [string]$Template,
        [System.Collections.Generic.List[object]]$SummaryItems
    )

    if ($Entries.Count -eq 0) {
        Write-Log "No TCG entries found in Excel file" -Level WARNING
        return
    }

    Write-Log "Found $($Entries.Count) TCG entries to process"

    $ProcessedCount = 0
    $TotalCount     = $Entries.Count

    foreach ($Entry in $Entries) {
        $OldFile = Find-MatchingFile -DownloadsFolder $DownloadsFolder -AccountNumber $Entry.AccountNumber
        if (-not $OldFile) { continue }

        $ProcessedCount++

        Write-Log "========================================"
        Write-Log "Matched $($Entry.AccountNumber) to $(Split-Path $OldFile -Leaf)" -Level MATCH

        $SafeEmployerName = Sanitize-FileName -FileName $Entry.EmployerName
        $SafeAmount       = Sanitize-FileName -FileName $Entry.Amount

        $NewBase = $Template -replace '\[EMPLOYER_NAME\]', $SafeEmployerName
        $NewBase = $NewBase  -replace '\[AMOUNT\]',        $SafeAmount

        $NewName            = Get-UniqueFileName -Directory $DownloadsFolder -BaseFileName $NewBase
        $NewPathInDownloads = Join-Path $DownloadsFolder $NewName
        
        try {
            Write-Log "Renaming: $(Split-Path $OldFile -Leaf) -> $NewName"
            Rename-Item -Path $OldFile -NewName $NewName -ErrorAction Stop
            Write-Log "Successfully renamed to $NewName" -Level INFO
            
            $FinalDestination = Join-Path $RemittanceFolder $NewName
            $Counter = 2
            while (Test-Path $FinalDestination) {
                if ($NewBase -match '^(JNT\s+\S+)(\s+.*)$') {
                    $DupName = "$($matches[1]) ($Counter)$($matches[2]).txt"
                } else {
                    $DupName = "$NewBase ($Counter).txt"
                }
                $FinalDestination = Join-Path $RemittanceFolder $DupName
                $Counter++
            }
            
            Write-Log "Moving file to Input_TPA: $(Split-Path $FinalDestination -Leaf)"
            Move-Item -Path $NewPathInDownloads -Destination $FinalDestination -ErrorAction Stop
            Write-Log "Successfully moved $($Entry.EmployerName) $($Entry.Amount) to Input_TPA - $ProcessedCount/$TotalCount" -Level SUCCESS

            $SummaryItems.Add([PSCustomObject]@{
                Company     = "TCG"
                DisplayName = $Entry.EmployerName
                FileName    = (Split-Path $FinalDestination -Leaf)
                Amount      = [decimal]$Entry.Amount
                AmountStr   = $Entry.Amount
            })
        }
        catch {
            Write-Log "Failed to process $(Split-Path $OldFile -Leaf): $($_.Exception.Message)" -Level ERROR
        }
    }
}

# ================= MAIN LOGIC (FIXED-FOLDER) =================

function Process-FixedFolders {
    param(
        [array]$Entries,
        [System.Collections.Generic.List[object]]$SummaryItems
    )
    
    Write-Log ""
    Write-Log "========================================"
    Write-Log "FIXED-FOLDER PROCESSING STARTED"
    Write-Log "========================================"
    Write-Log ""
    
    if ($Entries.Count -eq 0) {
        Write-Log "No Fixed-Folder entries found in Excel file" -Level WARNING
        return
    }
    
    Write-Log "Found $($Entries.Count) Fixed-Folder entries to process"
    Write-Log ""
    
    foreach ($Entry in $Entries) {
        switch ($Entry.Company) {
            "NBS"          { Process-NBS         -Entry $Entry -SummaryItems $SummaryItems }
            "Mid Atlantic" { Process-MidAtlantic  -Entry $Entry -SummaryItems $SummaryItems }
            "Penserv"      { Process-Penserv      -Entry $Entry -SummaryItems $SummaryItems }
        }
    }
    
    Write-Log ""
    Write-Log "========================================"
    Write-Log "FIXED-FOLDER PROCESSING COMPLETE"
    Write-Log "========================================"
}

# ================= SUMMARY PRINTER =================

function Print-Summary {
    param([System.Collections.Generic.List[object]]$SummaryItems)

    Write-Log ""
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  PROCESSING SUMMARY (sorted by amount)" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan

    if ($SummaryItems.Count -eq 0) {
        Write-Host "  No items were successfully processed." -ForegroundColor Yellow
        Write-Host "========================================" -ForegroundColor Cyan
        Add-Content -Path $LogFile -Value "SUMMARY: No items were successfully processed."
        return
    }

    $Sorted = $SummaryItems | Sort-Object Amount

    $MaxCompany = ($Sorted | ForEach-Object { $_.Company.Length }     | Measure-Object -Maximum).Maximum
    $MaxName    = ($Sorted | ForEach-Object { $_.DisplayName.Length } | Measure-Object -Maximum).Maximum
    $MaxAmount  = ($Sorted | ForEach-Object { $_.AmountStr.Length }   | Measure-Object -Maximum).Maximum

    # Column widths (minimum widths so headers fit)
    $ColCompany = [Math]::Max($MaxCompany, 12)
    $ColName    = [Math]::Max($MaxName,    30)
    $ColAmount  = [Math]::Max($MaxAmount,  10)

    $Header = "  {0,-$ColCompany}  {1,-$ColName}  {2,$ColAmount}  {3}" -f "Company", "Employer / Name", "Amount", "File"
    $Divider = "  " + ("-" * ($ColCompany + $ColName + $ColAmount + 40))

    Write-Host $Header  -ForegroundColor Cyan
    Write-Host $Divider -ForegroundColor Cyan
    Add-Content -Path $LogFile -Value ""
    Add-Content -Path $LogFile -Value "========================================"
    Add-Content -Path $LogFile -Value "PROCESSING SUMMARY (sorted by amount)"
    Add-Content -Path $LogFile -Value "========================================"
    Add-Content -Path $LogFile -Value $Header

    foreach ($Item in $Sorted) {
        $Line = "  {0,-$ColCompany}  {1,-$ColName}  {2,$ColAmount}  {3}" -f `
            $Item.Company, $Item.DisplayName, $Item.AmountStr, $Item.FileName
        Write-Host $Line -ForegroundColor Green
        Add-Content -Path $LogFile -Value $Line
    }

    Write-Host $Divider -ForegroundColor Cyan
    $TotalLine = "  {0,-$ColCompany}  {1}" -f "TOTAL", "($($Sorted.Count) items processed)"
    Write-Host $TotalLine -ForegroundColor Cyan
    Add-Content -Path $LogFile -Value $TotalLine
    Write-Host "========================================" -ForegroundColor Cyan
    Add-Content -Path $LogFile -Value "========================================"
}

# ================= EXECUTION =================

$ExcelFile        = $ExcelFile.Trim('"',' ')
$RemittanceFolder = $RemittanceFolder.Trim('"',' ')
$Downloads        = $Downloads.Trim('"',' ')

$LogFile = Join-Path $Downloads "TCG_Rename_$(Get-Date -Format 'MM-dd-yyyy_HH-mm').log"

# Shared list that every processor appends successfully-moved items to
$SummaryItems = [System.Collections.Generic.List[object]]::new()

Write-Log "========================================"
Write-Log "TCG FILE RENAMER + FIXED-FOLDER PROCESSOR"
Write-Log "========================================"
Write-Log "Excel file: $ExcelFile"
Write-Log "Remittance folder: $RemittanceFolder"
Write-Log "Downloads folder: $Downloads"
Write-Log "Template: $FileNameTemplate"
Write-Log ""

try {
    # ---- Extract everything from the morning report in a single Excel open ----
    Write-Log "========================================"
    Write-Log "EXTRACTING ALL DATA FROM MORNING REPORT"
    Write-Log "========================================"
    Write-Log ""

    $AllData = Extract-AllData -ExcelPath $ExcelFile

    Write-Log ""

    # ---- TCG ----
    Write-Log "========================================"
    Write-Log "TCG PROCESSING STARTED"
    Write-Log "========================================"
    Write-Log ""

    Rename-Files -Entries $AllData.TCG `
                 -DownloadsFolder $Downloads `
                 -RemittanceFolder $RemittanceFolder `
                 -Template $FileNameTemplate `
                 -SummaryItems $SummaryItems

    Write-Log ""
    Write-Log "========================================"
    Write-Log "TCG PROCESSING COMPLETE"
    Write-Log "========================================"
    Write-Log ""

    # ---- Fixed-Folder (NBS / Mid Atlantic / Penserv) ----
    Process-FixedFolders -Entries $AllData.FixedFolder -SummaryItems $SummaryItems

    # ---- SchoolsFirst (email download - prompt user if entries exist) ----
    if ($AllData.SchoolsFirst.Count -gt 0) {
        Write-Log ""

        $SFWarning = @"

========================================
  !! STOP - ACTION REQUIRED !!
========================================
  $($AllData.SchoolsFirst.Count) SchoolsFirst remittance(s) found.

  Please download the SchoolsFirst
  remittance file(s) from the
  DirectDeposit Email and save them
  to your Downloads folder.

  Return here and press Enter when done.
========================================

"@
        Write-Host $SFWarning -ForegroundColor Red
        Add-Content -Path $LogFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - WARNING - PAUSING: User prompted to download SchoolsFirst remittance from DirectDeposit Email"

        Read-Host "Press Enter once all SchoolsFirst remittance files have been downloaded to continue..."
        Write-Log "User confirmed SchoolsFirst downloads complete - proceeding with SchoolsFirst processing" -Level INFO

        Process-SchoolsFirst -Entries $AllData.SchoolsFirst `
                             -DownloadsFolder $Downloads `
                             -SummaryItems $SummaryItems
    } else {
        Write-Log "No SchoolsFirst entries found - skipping SchoolsFirst processing" -Level INFO
    }

    # ---- AFPlanServ (email download - prompt user if entries exist) ----
    if ($AllData.AFP.Count -gt 0) {
        Write-Log ""

        $AFPWarning = @"

========================================
  !! STOP - ACTION REQUIRED !!
========================================
  $($AllData.AFP.Count) AFPlanServ remittance(s) found.

  Please download the AFPlanServ
  remittance file(s) from the email
  and save them to your Downloads folder.

  Return here and press Enter when done.
========================================

"@
        Write-Host $AFPWarning -ForegroundColor Red
        Add-Content -Path $LogFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - WARNING - PAUSING: User prompted to download AFPlanServ remittance from email"

        Read-Host "Press Enter once all AFPlanServ remittance files have been downloaded to continue..."
        Write-Log "User confirmed AFPlanServ downloads complete - proceeding with AFPlanServ processing" -Level INFO

        Process-AFPlanServ -Entries $AllData.AFP `
                           -DownloadsFolder $Downloads `
                           -SummaryItems $SummaryItems
    } else {
        Write-Log "No AFPlanServ entries found - skipping AFPlanServ processing" -Level INFO
    }

    # ---- OMNI (website download - prompt user only if entries exist) ----
    if ($AllData.OMNI.Count -gt 0) {
        Write-Log ""

        $RedWarning = @"

========================================
  !! STOP - ACTION REQUIRED !!
========================================
  $($AllData.OMNI.Count) OMNI remittance(s) found.

  OMNI remittance files must be
  downloaded BEFORE continuing.

  Opening https://www.omni403b.com/ now.

  1. Log in and download ALL remittance
     files for today to your Downloads
     folder.
  2. Return here and press Enter when
     done to continue processing.
========================================

"@
        Write-Host $RedWarning -ForegroundColor Red
        Add-Content -Path $LogFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - WARNING - PAUSING: User prompted to download OMNI remittance files from https://www.omni403b.com/"

        try {
            Start-Process "https://www.omni403b.com/"
        }
        catch {
            Write-Host "Could not auto-open browser. Please navigate to: https://www.omni403b.com/" -ForegroundColor Red
        }

        Read-Host "Press Enter once all OMNI remittance files have been downloaded to continue..."
        Write-Log "User confirmed OMNI downloads complete - proceeding with OMNI processing" -Level INFO

        Process-OMNI -Entries $AllData.OMNI `
                     -DownloadsFolder $Downloads `
                     -SummaryItems $SummaryItems
    } else {
        Write-Log "No OMNI entries found - skipping OMNI processing" -Level INFO
    }

    # ---- Alta Montclair (email download - prompt user if entries exist) ----
    if ($AllData.AM.Count -gt 0) {
        Write-Log ""

        $AMWarning = @"

========================================
  !! STOP - ACTION REQUIRED !!
========================================
  $($AllData.AM.Count) Alta Montclair remittance(s) found.

  Please download the Alta Montclair
  remittance file(s) from the email
  and save them to your Downloads folder.

  Note: Alta Montclair files may show
  security warnings when opened in Excel
  manually - that's expected.

  Return here and press Enter when done.
========================================

"@
        Write-Host $AMWarning -ForegroundColor Red
        Add-Content -Path $LogFile -Value "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') - WARNING - PAUSING: User prompted to download Alta Montclair remittance from email"

        Read-Host "Press Enter once all Alta Montclair remittance files have been downloaded to continue..."
        Write-Log "User confirmed Alta Montclair downloads complete - proceeding with Alta Montclair processing" -Level INFO

        Process-AltaMontclair -Entries $AllData.AM `
                              -DownloadsFolder $Downloads `
                              -SummaryItems $SummaryItems
    } else {
        Write-Log "No Alta Montclair entries found - skipping Alta Montclair processing" -Level INFO
    }

    # ---- Final summary ----
    Print-Summary -SummaryItems $SummaryItems

    Write-Log ""
    Write-Log "========================================"
    Write-Log "ALL PROCESSING COMPLETE"
    Write-Log "Remember to double check log to verify correctness"
    Write-Log "========================================"

} catch {
    Write-Log "Fatal error: $($_.Exception.Message)" -Level ERROR
    Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level ERROR
}

Write-Host ""
Write-Host "Log file saved to: $LogFile" -ForegroundColor Cyan
Write-Host ""
Read-Host "Press Enter to exit"
