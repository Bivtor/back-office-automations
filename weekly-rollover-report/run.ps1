# Weekly Rollover Report - Generation Script

param(
    [Parameter(Mandatory=$true)]
    [string]$RawInputFile,

    [Parameter(Mandatory=$false)]
    [string]$OldReportFile = ""
)

# Clean up parameters
$RawInputFile  = $RawInputFile.Trim('"', ' ')
$OldReportFile = $OldReportFile.Trim('"', ' ')

# Get script directory
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path


# ================= CONFIGURATION LOADING =================
$ConfigPath = Join-Path $ScriptDir "config.json"
if (-not (Test-Path $ConfigPath)) {
    Write-Error "Configuration file not found at $ConfigPath. Please ensure config.json is present."
    Read-Host "Press Enter to exit"
    exit 1
}
$Config = Get-Content $ConfigPath | ConvertFrom-Json

# Extract configuration values
$OutputBasePath  = $Config.outputBasePath
$HeaderColor     = $Config.formatting.headerColor
$HeaderFontColor = $Config.formatting.headerFontColor

# Column order for all sheets
$ColumnOrder = $Config.columnOrder

# RO Report filters
$ROReportExcludeCompanies = $Config.sheets.roReport.filters.excludeSourceCompanies

# No TPA filters
$NoTPABusinessDayThreshold = [int]$Config.sheets.noTPA.filters.tpaRequestedBusinessDaysThreshold

# Holiday config
$HolidayApiUrl         = $Config.holidays.apiUrl
$HolidayApiTimeout     = [int]$Config.holidays.apiTimeoutSeconds
$HolidayFallbackList   = $Config.holidays.fallbackHolidays


# ================= OUTPUT FOLDER SETUP =================
$Timestamp         = Get-Date -Format 'M.d.yyyy'
$TimeComponent     = Get-Date -Format 'h-mmtt'
$OutputFolderName  = "$Timestamp Weekly Rollover Report $TimeComponent"

# Allow config override for output base path, fallback to input file directory
if (-not [string]::IsNullOrWhiteSpace($OutputBasePath) -and (Test-Path $OutputBasePath)) {
    $OutputFolder = Join-Path $OutputBasePath $OutputFolderName
} else {
    $InputFileDirectory = Split-Path -Parent $RawInputFile
    $OutputFolder = Join-Path $InputFileDirectory $OutputFolderName
}

if (-not (Test-Path $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory | Out-Null
}

$OutputFileName = "$OutputFolderName.xlsx"
$OutputFilePath = Join-Path $OutputFolder $OutputFileName
$LogFile        = Join-Path $OutputFolder "$OutputFolderName.log"


# ================= LOGGING =================
function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARNING','ERROR','SUCCESS')]
        [string]$Level = 'INFO'
    )

    $LogTimestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $Line = "$LogTimestamp - $Level - $Message"

    switch ($Level) {
        'INFO'    { Write-Host $Line -ForegroundColor White }
        'WARNING' { Write-Host $Line -ForegroundColor Yellow }
        'ERROR'   { Write-Host $Line -ForegroundColor Red }
        'SUCCESS' { Write-Host $Line -ForegroundColor Green }
    }

    Add-Content -Path $LogFile -Value $Line
}

Write-Log "========================================"
Write-Log "Weekly Rollover Report Processing Started"
Write-Log "========================================"
Write-Log "Raw Input File : $RawInputFile"
Write-Log "Output Folder  : $OutputFolder"
Write-Log "Output File    : $OutputFilePath"
Write-Log "Log File       : $LogFile"
Write-Log "Config File    : $ConfigPath"
Write-Log ""
Write-Log "Configuration loaded successfully" -Level SUCCESS
Write-Log ""


# ================= IMPORTEXCEL CHECK =================
Write-Log "Checking ImportExcel module..."
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Log "ImportExcel module not found. Attempting to install..." -Level WARNING
    try {
        Install-Module -Name ImportExcel -Scope CurrentUser -Force -AllowClobber
        Write-Log "ImportExcel installed successfully" -Level SUCCESS
    }
    catch {
        Write-Log "Failed to install ImportExcel: $($_.Exception.Message)" -Level ERROR
        Write-Log "Please run: Install-Module -Name ImportExcel -Scope CurrentUser -Force" -Level ERROR
        Read-Host "Press Enter to exit"
        exit 1
    }
}
Import-Module ImportExcel -Force
Write-Log "ImportExcel module loaded" -Level SUCCESS
Write-Log ""


# ================= HELPER: Parse hex color to ARGB for ImportExcel =================
function ConvertTo-ExcelColor {
    param([string]$HexColor)
    # ImportExcel / ClosedXML use System.Drawing.Color
    return [System.Drawing.ColorTranslator]::FromHtml($HexColor)
}


# ================= HELPER: Reposition columns =================
# Returns a new ordered list of PSCustomObjects with only the defined columns (in order),
# ignoring any columns not in the target list, and appending any extra columns at the end.
function Invoke-ColumnReorder {
    param(
        [System.Collections.Generic.List[object]]$Rows,
        [array]$DesiredOrder
    )

    if ($Rows.Count -eq 0) { return $Rows }

    # Get all property names from the first row
    $AllColumns = $Rows[0].PSObject.Properties.Name

    # Build final column list: desired columns that exist, then any remaining extras
    $ExistingDesired = $DesiredOrder | Where-Object { $AllColumns -contains $_ }
    $Extras          = $AllColumns   | Where-Object { $DesiredOrder -notcontains $_ }
    $FinalOrder      = @($ExistingDesired) + @($Extras)

    $Reordered = [System.Collections.Generic.List[object]]::new()
    foreach ($Row in $Rows) {
        $Obj = [ordered]@{}
        foreach ($Col in $FinalOrder) {
            $Obj[$Col] = $Row.$Col
        }
        $Reordered.Add([PSCustomObject]$Obj)
    }

    return $Reordered
}


# ================= HELPER: Remove rows by Source Company partial match =================
function Remove-ExcludedCompanies {
    param(
        [array]$Rows,
        [array]$ExcludeList
    )

    if (-not $ExcludeList -or $ExcludeList.Count -eq 0) {
        Write-Log "  No company exclusions configured - skipping"
        return $Rows
    }

    $Before = $Rows.Count
    $Filtered = $Rows | Where-Object {
        $SourceCompany = "$($_.'Source Company')".Trim()
        $exclude = $false
        foreach ($Term in $ExcludeList) {
            if ($SourceCompany -match [regex]::Escape($Term)) {
                $exclude = $true
                break
            }
        }
        -not $exclude
    }

    $After = @($Filtered).Count
    Write-Log "  Excluded companies: removed $($Before - $After) rows (from $Before to $After)"
    return $Filtered
}


# ================= HELPER: Remove duplicate Rollover IDs =================
function Remove-DuplicateRollovers {
    param(
        [array]$Rows
    )

    $Seen    = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $Before  = $Rows.Count
    $Deduped = $Rows | Where-Object {
        $RID = "$($_.'Rollover ID')".Trim()
        if ([string]::IsNullOrWhiteSpace($RID)) { return $true }  # keep rows with no ID
        $Seen.Add($RID)  # returns $false if already present, which Where-Object treats as exclude
    }

    $After = @($Deduped).Count
    Write-Log "  Duplicate Rollover IDs: removed $($Before - $After) rows (from $Before to $After)"
    return $Deduped
}


# ================= HELPER: Remove rows where OES column is blank =================
function Remove-BlankOES {
    param(
        [array]$Rows
    )

    if ($Rows.Count -eq 0) { return $Rows }

    # Find OES column case-insensitively (exact match)
    $OESCol = $Rows[0].PSObject.Properties.Name | Where-Object { $_ -ieq 'OES' } | Select-Object -First 1

    if (-not $OESCol) {
        Write-Log "  OES column not found - skipping blank OES filter" -Level WARNING
        return $Rows
    }

    $Before   = $Rows.Count
    $Filtered = $Rows | Where-Object { -not [string]::IsNullOrWhiteSpace($_.$OESCol) }
    $After    = @($Filtered).Count
    Write-Log "  Blank OES removed: $($Before - $After) rows (from $Before to $After)"
    return $Filtered
}


# ================= HELPER: Read Category + Notes from old report's No LOA sheet =================
function Get-OldNoLOAData {
    param([string]$OldReportPath)

    # Returns a hashtable of RolloverID -> @{ Category = "..."; Notes = "..." }
    $DataMap = @{}

    if ([string]::IsNullOrWhiteSpace($OldReportPath) -or -not (Test-Path $OldReportPath)) {
        Write-Log "  No old report available - skipping carry-over"
        return $DataMap
    }

    Write-Log "  Reading old report No LOA sheet: $OldReportPath"

    try {
        $OldRows = $null
        try {
            $OldRows = Import-Excel -Path $OldReportPath -WorksheetName "No LOA"
        } catch {
            Write-Log "    'No LOA' sheet not found in old report - skipping" -Level WARNING
            return $DataMap
        }

        if (-not $OldRows -or $OldRows.Count -eq 0) {
            Write-Log "    Old No LOA sheet is empty - skipping"
            return $DataMap
        }

        # Find column names case-insensitively
        $RIDCol      = $OldRows[0].PSObject.Properties.Name | Where-Object { $_ -match '(?i)rollover\s*id' } | Select-Object -First 1
        $NotesCol    = $OldRows[0].PSObject.Properties.Name | Where-Object { $_ -match '(?i)^notes$' }       | Select-Object -First 1
        $CategoryCol = $OldRows[0].PSObject.Properties.Name | Where-Object { $_ -match '(?i)^category$' }   | Select-Object -First 1

        if (-not $RIDCol) {
            Write-Log "    'Rollover ID' column not found in old No LOA sheet" -Level WARNING
            return $DataMap
        }

        Write-Log "    Rollover ID col : '$RIDCol'"
        Write-Log "    Category col    : $(if ($CategoryCol) { '$CategoryCol' } else { 'NOT FOUND' })"
        Write-Log "    Notes col       : $(if ($NotesCol) { '$NotesCol' } else { 'NOT FOUND' })"

        foreach ($Row in $OldRows) {
            $RID = "$($Row.$RIDCol)".Trim()
            if ([string]::IsNullOrWhiteSpace($RID)) { continue }

            $Category = if ($CategoryCol) { "$($Row.$CategoryCol)".Trim() } else { "" }
            $Notes    = if ($NotesCol)    { "$($Row.$NotesCol)".Trim()    } else { "" }

            # Store for any row with a Rollover ID, even if both fields are blank
            # (we still want to carry the row's category even if notes is empty)
            if (-not [string]::IsNullOrWhiteSpace($Category) -or -not [string]::IsNullOrWhiteSpace($Notes)) {
                $DataMap[$RID] = @{ Category = $Category; Notes = $Notes }
            }
        }

        Write-Log "    Loaded $($DataMap.Count) carry-over entries (Category and/or Notes)" -Level SUCCESS
    }
    catch {
        Write-Log "    Error reading old report: $($_.Exception.Message)" -Level ERROR
    }

    return $DataMap
}


# ================= HELPER: Read Within Timing from old report's No TPA sheet =================
function Get-OldNoTPAData {
    param([string]$OldReportPath)

    # Returns a hashtable of RolloverID -> @{ WithinTiming = "..." }
    # NOTE: Notes column exists on No TPA tab but is never carried over from old reports
    $DataMap = @{}

    if ([string]::IsNullOrWhiteSpace($OldReportPath) -or -not (Test-Path $OldReportPath)) {
        Write-Log "  No old report available - skipping No TPA carry-over"
        return $DataMap
    }

    Write-Log "  Reading old report No TPA sheet: $OldReportPath"

    try {
        $OldRows = $null
        try {
            $OldRows = Import-Excel -Path $OldReportPath -WorksheetName "No TPA"
        } catch {
            Write-Log "    'No TPA' sheet not found in old report - skipping" -Level WARNING
            return $DataMap
        }

        if (-not $OldRows -or $OldRows.Count -eq 0) {
            Write-Log "    Old No TPA sheet is empty - skipping"
            return $DataMap
        }

        # Find column names case-insensitively
        $RIDCol          = $OldRows[0].PSObject.Properties.Name | Where-Object { $_ -match '(?i)rollover\s*id' }     | Select-Object -First 1
        $WithinTimingCol = $OldRows[0].PSObject.Properties.Name | Where-Object { $_ -match '(?i)within\s*timing' }   | Select-Object -First 1

        if (-not $RIDCol) {
            Write-Log "    'Rollover ID' column not found in old No TPA sheet" -Level WARNING
            return $DataMap
        }

        Write-Log "    Rollover ID col    : '$RIDCol'"
        Write-Log "    Within Timing col  : $(if ($WithinTimingCol) { $WithinTimingCol } else { 'NOT FOUND' })"

        foreach ($Row in $OldRows) {
            $RID = "$($Row.$RIDCol)".Trim()
            if ([string]::IsNullOrWhiteSpace($RID)) { continue }

            $WithinTiming = if ($WithinTimingCol) { "$($Row.$WithinTimingCol)".Trim() } else { "" }

            if (-not [string]::IsNullOrWhiteSpace($WithinTiming)) {
                $DataMap[$RID] = @{ WithinTiming = $WithinTiming }
            }
        }

        Write-Log "    Loaded $($DataMap.Count) carry-over entries from old No TPA sheet" -Level SUCCESS
    }
    catch {
        Write-Log "    Error reading old No TPA sheet: $($_.Exception.Message)" -Level ERROR
    }

    return $DataMap
}


# ================= STEP 1: COPY INPUT FILE & LOAD DATA =================
Write-Log "========================================"
Write-Log "STEP 1: Copying Input File & Loading Data"
Write-Log "========================================"
Write-Log ""

if (-not (Test-Path $RawInputFile)) {
    Write-Log "ERROR: Input file not found: $RawInputFile" -Level ERROR
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Log "Copying input file to output folder..."
Copy-Item -Path $RawInputFile -Destination $OutputFilePath -Force
Write-Log "Copied to: $OutputFilePath" -Level SUCCESS
Write-Log ""

# ================= STEP 2: CHECK AND CLEAN TITLE ROW =================
Write-Log "========================================"
Write-Log "STEP 2: Checking for Title Row"
Write-Log "========================================"
Write-Log ""

# Open the file directly to check A1 before doing a full import.
# Import-Excel always uses row 1 as headers - so if it's "Rollover Report" we
# delete that row in the file first, then import cleanly.
Write-Log "Checking cell A1..."
$CheckPkg = Open-ExcelPackage -Path $OutputFilePath
$CheckWS  = $CheckPkg.Workbook.Worksheets | Select-Object -First 1
$SheetName = $CheckWS.Name
$CellA1   = "$($CheckWS.Cells[1,1].Value)".Trim().ToLower()
Write-Log "  Sheet name : '$SheetName'"
Write-Log "  Cell A1    : '$CellA1'"

if ($CellA1 -eq "rollover report") {
    Write-Log "  Title row detected - deleting row 1 from file..."
    $CheckWS.DeleteRow(1)
    $CheckPkg.Save()
    Write-Log "  Row 1 deleted and file saved" -Level SUCCESS
} else {
    Write-Log "  Row 1 looks like real headers - no removal needed"
}
$CheckPkg.Dispose()
Write-Log ""

# Now do the real import with correct headers in row 1
Write-Log "Importing data from file (sheet: '$SheetName')..."
try {
    $RawData = Import-Excel -Path $OutputFilePath -WorksheetName $SheetName
    Write-Log "Loaded $($RawData.Count) rows" -Level SUCCESS
    Write-Log "  Columns: $($RawData[0].PSObject.Properties.Name -join ', ')"
}
catch {
    Write-Log "ERROR: Could not load input data: $($_.Exception.Message)" -Level ERROR
    Read-Host "Press Enter to exit"
    exit 1
}

# ================= STEP 3: WRITE DATA SHEET =================
Write-Log "========================================"
Write-Log "STEP 3: Writing 'Data' Sheet"
Write-Log "========================================"
Write-Log ""

Write-Log "Writing data to 'Data' sheet..."
try {
    # Write raw data as-is, no column reorder on Data sheet
    Remove-Item -Path $OutputFilePath -Force

    $RawData | Export-Excel `
        -Path $OutputFilePath `
        -WorksheetName "Data" `
        -AutoFilter `
        -FreezeTopRow

    Write-Log "Data sheet written successfully ($($RawData.Count) rows)" -Level SUCCESS
}
catch {
    Write-Log "ERROR writing Data sheet: $($_.Exception.Message)" -Level ERROR
    Read-Host "Press Enter to exit"
    exit 1
}
Write-Log ""


# ================= STEPS 4-7: BUILD ALL SHEET DATA IN MEMORY =================
Write-Log "========================================"
Write-Log "STEPS 4-7: Building Sheet Data In Memory"
Write-Log "========================================"
Write-Log ""

# --- RO Report: filter internals, dedupe, remove blank OES, reorder columns ---
Write-Log "STEP 4: Building RO Report data..."
Write-Log "  Removing excluded source companies..."
$ROReportFiltered = Remove-ExcludedCompanies -Rows $RawData -ExcludeList $ROReportExcludeCompanies
Write-Log "  Removing duplicate Rollover IDs..."
$ROReportDeduped  = Remove-DuplicateRollovers -Rows $ROReportFiltered
Write-Log "  Removing rows with blank OES..."
$ROReportDeduped  = Remove-BlankOES -Rows $ROReportDeduped
$ROReportData = Invoke-ColumnReorder -Rows ([System.Collections.Generic.List[object]]$ROReportDeduped) -DesiredOrder $ColumnOrder
Write-Log "  RO Report final: $($ROReportData.Count) rows" -Level SUCCESS

# --- No LOA: deep copy of RO Report, filtered to rows missing LOA Date, TPA Auth, and TPA Approval ---
Write-Log "STEP 5: Building No LOA data..."

# Deep copy from the already-filtered/deduped RO Report base (before column reorder)
$NoLOABase = $ROReportDeduped | ForEach-Object {
    $NewObj = [ordered]@{}
    foreach ($Prop in $_.PSObject.Properties) { $NewObj[$Prop.Name] = $Prop.Value }
    [PSCustomObject]$NewObj
}

# Find column names case-insensitively
$NoLOA_LOACol     = $NoLOABase[0].PSObject.Properties.Name | Where-Object { $_ -match '(?i)^loa\s*date$' }                    | Select-Object -First 1
$NoLOA_TPAAuthCol = $NoLOABase[0].PSObject.Properties.Name | Where-Object { $_ -match '(?i)tpa\s*authorization\s*requested' } | Select-Object -First 1
$NoLOA_TPAApprCol = $NoLOABase[0].PSObject.Properties.Name | Where-Object { $_ -match '(?i)tpa\s*approval' }                  | Select-Object -First 1

Write-Log "  LOA Date col      : '$NoLOA_LOACol'"
Write-Log "  TPA Auth col      : '$NoLOA_TPAAuthCol'"
Write-Log "  TPA Approval col  : '$NoLOA_TPAApprCol'"

# Filter: LOA Date, TPA Auth, and TPA Approval must all be blank
$NoLOAFiltered = $NoLOABase | Where-Object {
    ($NoLOA_LOACol     -and [string]::IsNullOrWhiteSpace($_.$NoLOA_LOACol)) -and
    ($NoLOA_TPAAuthCol -and [string]::IsNullOrWhiteSpace($_.$NoLOA_TPAAuthCol)) -and
    ($NoLOA_TPAApprCol -and [string]::IsNullOrWhiteSpace($_.$NoLOA_TPAApprCol))
}
Write-Log "  Rows after LOA/TPA filter: $(@($NoLOAFiltered).Count)"
# Note: blank OES rows already removed in STEP 4 from the base data

# Add Category and Notes columns as first two columns (blank by default)
$NoLOAWithNotes = @($NoLOAFiltered) | ForEach-Object {
    $NewObj = [ordered]@{}
    $NewObj["Category"] = ""
    $NewObj["Notes"]    = ""
    foreach ($Prop in $_.PSObject.Properties) { $NewObj[$Prop.Name] = $Prop.Value }
    [PSCustomObject]$NewObj
}

# Carry over Category and Notes from old report where Rollover ID matches
Write-Log "  Loading Category and Notes from old report..."
$OldNoLOAData = Get-OldNoLOAData -OldReportPath $OldReportFile
$CarryCount = 0

if ($OldNoLOAData.Count -gt 0) {
    $NoLOAWithNotes = $NoLOAWithNotes | ForEach-Object {
        $RID = "$($_.'Rollover ID')".Trim()
        if (-not [string]::IsNullOrWhiteSpace($RID) -and $OldNoLOAData.ContainsKey($RID)) {
            $OldEntry = $OldNoLOAData[$RID]

            # Carry over Category as-is
            if (-not [string]::IsNullOrWhiteSpace($OldEntry.Category)) {
                $_.Category = $OldEntry.Category
            }

            # Carry over Notes as-is (no prefix added)
            if (-not [string]::IsNullOrWhiteSpace($OldEntry.Notes)) {
                # Strip any legacy "From Old Report: " prefix that may exist in older reports
                $CleanedNotes = $OldEntry.Notes -replace '^From Old Report:\s*', ''
                $_.Notes = $CleanedNotes
            }

            $CarryCount++
        }
        $_
    }
    Write-Log "  Carried over data for $CarryCount matching Rollover IDs" -Level SUCCESS
}

# Sort No LOA by Category: non-blank first (alphabetically), blanks at the bottom
Write-Log "  Sorting No LOA by Category (non-blank first, blanks last)..."
$NoLOASorted = $NoLOAWithNotes | Sort-Object {
    $cat = "$($_.Category)".Trim()
    if ([string]::IsNullOrWhiteSpace($cat)) { "ZZZZZ" } else { $cat }
}

# Column order for No LOA: Category, Notes, then the standard columns
$NoLOAColumnOrder = @("Category", "Notes") + $ColumnOrder
$NoLOAData = Invoke-ColumnReorder -Rows ([System.Collections.Generic.List[object]]$NoLOASorted) -DesiredOrder $NoLOAColumnOrder
Write-Log "  No LOA: $($NoLOAData.Count) rows" -Level SUCCESS

# --- No TPA: overdue TPA requests (requested > N business days ago, not yet approved) ---
Write-Log "STEP 6: Building No TPA data..."

# Load US public holidays - try API first, fall back to config list
function Get-USHolidays {
    param(
        [string]$ApiUrl,
        [int]$TimeoutSec,
        [array]$FallbackList,
        [int[]]$Years
    )

    $HolidaySet = [System.Collections.Generic.HashSet[string]]::new()

    # Try API for each year needed
    $ApiSuccess = $false
    foreach ($Year in $Years) {
        $Url = $ApiUrl -replace '\{year\}', $Year
        try {
            $Response = Invoke-RestMethod -Uri $Url -TimeoutSec $TimeoutSec -ErrorAction Stop
            foreach ($Holiday in $Response) {
                $HolidaySet.Add($Holiday.date) | Out-Null
            }
            $ApiSuccess = $true
            Write-Log "    Loaded $($Response.Count) holidays for $Year from API"
        } catch {
            Write-Log "    API unavailable for $Year ($($_.Exception.Message)) - will use fallback" -Level WARNING
            $ApiSuccess = $false
            break
        }
    }

    # Fall back to config list if API failed
    if (-not $ApiSuccess) {
        Write-Log "    Using hardcoded holiday fallback list ($($FallbackList.Count) dates)" -Level WARNING
        foreach ($D in $FallbackList) {
            $HolidaySet.Add($D) | Out-Null
        }
    }

    return $HolidaySet
}

function Get-BusinessDateOffset {
    param(
        [datetime]$StartDate,
        [int]$OffsetDays,
        [System.Collections.Generic.HashSet[string]]$Holidays
    )
    $Current   = $StartDate
    $Remaining = [Math]::Abs($OffsetDays)
    $Direction = if ($OffsetDays -lt 0) { -1 } else { 1 }
    while ($Remaining -gt 0) {
        $Current = $Current.AddDays($Direction)
        $DateStr = $Current.ToString('yyyy-MM-dd')
        $IsWeekend = $Current.DayOfWeek -eq [System.DayOfWeek]::Saturday -or
                     $Current.DayOfWeek -eq [System.DayOfWeek]::Sunday
        $IsHoliday = $Holidays -and $Holidays.Contains($DateStr)
        if (-not $IsWeekend -and -not $IsHoliday) {
            $Remaining--
        }
    }
    return $Current
}

$Today      = (Get-Date).Date
$YearsNeeded = @($Today.Year)
if ($Today.Month -le 2) { $YearsNeeded += ($Today.Year - 1) }  # catch prior year holidays near Jan

Write-Log "  Fetching US holiday calendar..."
$USHolidays = Get-USHolidays -ApiUrl $HolidayApiUrl -TimeoutSec $HolidayApiTimeout -FallbackList $HolidayFallbackList -Years $YearsNeeded

$TwoBizDaysPrior = Get-BusinessDateOffset -StartDate $Today -OffsetDays (-$NoTPABusinessDayThreshold) -Holidays $USHolidays
Write-Log "  Today                    : $($Today.ToString('MM/dd/yyyy'))"
Write-Log "  Cutoff ($NoTPABusinessDayThreshold biz days prior) : $($TwoBizDaysPrior.ToString('MM/dd/yyyy'))"

# Deep copy from ROReportDeduped (same base as No LOA)
$NoTPABase = $ROReportDeduped | ForEach-Object {
    $NewObj = [ordered]@{}
    foreach ($Prop in $_.PSObject.Properties) { $NewObj[$Prop.Name] = $Prop.Value }
    [PSCustomObject]$NewObj
}

# Find column names case-insensitively
$NoTPA_LOACol     = $NoTPABase[0].PSObject.Properties.Name | Where-Object { $_ -match '(?i)^loa\s*date$' }                | Select-Object -First 1
$NoTPA_TPAAuthCol = $NoTPABase[0].PSObject.Properties.Name | Where-Object { $_ -match '(?i)tpa\s*authorization\s*requested' } | Select-Object -First 1
$NoTPA_TPAApprCol = $NoTPABase[0].PSObject.Properties.Name | Where-Object { $_ -match '(?i)tpa\s*approval' }               | Select-Object -First 1

Write-Log "  LOA Date col      : '$NoTPA_LOACol'"
Write-Log "  TPA Auth col      : '$NoTPA_TPAAuthCol'"
Write-Log "  TPA Approval col  : '$NoTPA_TPAApprCol'"

if (-not $NoTPA_TPAAuthCol) {
    Write-Log "  'TPA Authorization Requested' column not found - sheet will be empty" -Level WARNING
    $NoTPAFiltered = @()
} else {
    $NoTPAFiltered = $NoTPABase | Where-Object {
        # LOA Date must be blank
        ($NoTPA_LOACol  -and [string]::IsNullOrWhiteSpace($_.$NoTPA_LOACol)) -and
        # TPA Approval must be blank
        ($NoTPA_TPAApprCol -and [string]::IsNullOrWhiteSpace($_.$NoTPA_TPAApprCol)) -and
        # TPA Auth Requested must exist (not blank)
        (-not [string]::IsNullOrWhiteSpace($_.$NoTPA_TPAAuthCol)) -and
        # TPA Auth Requested must be before the business day cutoff
        (& {
            try {
                $RequestedDate = [datetime]::Parse("$($_.$NoTPA_TPAAuthCol)".Trim())
                # Exclude future-dated entries as a safety guard
                $RequestedDate -lt $TwoBizDaysPrior -and $RequestedDate -le $Today
            } catch {
                $false  # If date can't be parsed, exclude the row
            }
        })
    }
    Write-Log "  Rows after filter: $(@($NoTPAFiltered).Count)"
}
# Note: blank OES rows already removed in STEP 4 from the base data

# Add Notes and Within Timing columns at the front (Notes left blank - not carried over)
$NoTPAWithCols = @($NoTPAFiltered) | ForEach-Object {
    $NewObj = [ordered]@{}
    $NewObj["Notes"]         = ""
    $NewObj["Within Timing"] = ""
    foreach ($Prop in $_.PSObject.Properties) { $NewObj[$Prop.Name] = $Prop.Value }
    [PSCustomObject]$NewObj
}

# Neither Notes nor Within Timing are carried over from the old report - both columns left blank
Write-Log "  No TPA carry-over skipped - Notes and Within Timing left blank intentionally"

$NoTPAColumnOrder = @("Notes", "Within Timing") + $ColumnOrder
$NoTPAData = Invoke-ColumnReorder -Rows ([System.Collections.Generic.List[object]]$NoTPAWithCols) -DesiredOrder $NoTPAColumnOrder
Write-Log "  No TPA: $($NoTPAData.Count) rows" -Level SUCCESS

Write-Log ""

# ================= STEP 7: WRITE ALL SHEETS WITH INLINE FORMATTING =================
Write-Log "========================================"
Write-Log "STEP 7: Writing All Sheets to File (with formatting)"
Write-Log "========================================"
Write-Log ""

# Helper scriptblock - applies bold headers, freeze, autofilter, autofit
# $FixedWidths is an optional hashtable of { "ColumnHeaderName" = width }
$FormatWS = {
    param($WS, $BgColor, $FontColor, $FixedWidths = @{}, $WrapColumns = @())
    $LastCol = $WS.Dimension.End.Column
    $LastRow = $WS.Dimension.End.Row

    for ($col = 1; $col -le $LastCol; $col++) {
        $Cell = $WS.Cells[1, $col]
        $Cell.Style.Font.Bold = $true
        $Cell.Style.Font.Color.SetColor($FontColor)
        $Cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $Cell.Style.Fill.BackgroundColor.SetColor($BgColor)
        $Cell.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
    }

    if ($LastRow -gt 1) {
        $WS.Cells[2, 1, $LastRow, $LastCol].Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
    }

    # Autofit all columns first
    for ($col = 1; $col -le $LastCol; $col++) {
        $WS.Column($col).AutoFit()
    }

    # Then apply any fixed-width overrides by header name
    if ($FixedWidths.Count -gt 0) {
        for ($col = 1; $col -le $LastCol; $col++) {
            $HeaderName = $WS.Cells[1, $col].Text
            if ($FixedWidths.ContainsKey($HeaderName)) {
                $WS.Column($col).Width = $FixedWidths[$HeaderName]
            }
        }
    }

    # Apply wrap text to specified columns (by header name), data rows only
    if ($WrapColumns.Count -gt 0) {
        for ($col = 1; $col -le $LastCol; $col++) {
            $HeaderName = $WS.Cells[1, $col].Text
            if ($WrapColumns -contains $HeaderName) {
                if ($LastRow -gt 1) {
                    $WS.Cells[2, $col, $LastRow, $col].Style.WrapText = $true
                }
            }
        }
    }

    $WS.View.FreezePanes(2, 1)
    $WS.Cells[1, 1, 1, $LastCol].AutoFilter = $true
}

$BgColor   = [System.Drawing.ColorTranslator]::FromHtml($HeaderColor)
$FontColor = [System.Drawing.ColorTranslator]::FromHtml($HeaderFontColor)

# Write Data sheet (already written, reopen once to format it)
Write-Log "Formatting Data sheet..."
try {
    $Pkg = Open-ExcelPackage -Path $OutputFilePath
    & $FormatWS $Pkg.Workbook.Worksheets["Data"] $BgColor $FontColor
    $Pkg.Save()
    $Pkg.Dispose()
    Write-Log "  Data sheet formatted" -Level SUCCESS
} catch {
    Write-Log "  Error formatting Data sheet: $($_.Exception.Message)" -Level ERROR
    try { $Pkg.Dispose() } catch {}
}

# Write RO Report with formatting via -PassThru
Write-Log "Writing and formatting RO Report sheet..."
$Pkg = $ROReportData | Export-Excel -Path $OutputFilePath -WorksheetName "RO Report" -ClearSheet -PassThru
& $FormatWS $Pkg.Workbook.Worksheets["RO Report"] $BgColor $FontColor
$Pkg.Save()
$Pkg.Dispose()
Write-Log "  RO Report written and formatted" -Level SUCCESS

# Write No LOA with formatting via -PassThru
Write-Log "Writing and formatting No LOA sheet..."
$NoLOAExportData = if ($NoLOAData.Count -gt 0) { $NoLOAData } else {
    # Build an empty placeholder row that includes Category and Notes up front,
    # then the standard column order - so headers always appear even when the sheet is empty
    $EmptyColOrder = @("Category", "Notes") + $ColumnOrder
    $EmptyRow = [ordered]@{}
    foreach ($Col in $EmptyColOrder) { $EmptyRow[$Col] = "" }
    @([PSCustomObject]$EmptyRow)
}
$Pkg = $NoLOAExportData | Export-Excel -Path $OutputFilePath -WorksheetName "No LOA" -ClearSheet -PassThru
& $FormatWS $Pkg.Workbook.Worksheets["No LOA"] $BgColor $FontColor @{ "Notes" = 50 } @("Notes")
$Pkg.Save()
$Pkg.Dispose()
Write-Log "  No LOA written and formatted" -Level SUCCESS

# Write No TPA with formatting via -PassThru
Write-Log "Writing and formatting No TPA sheet..."
$NoTPAExportData = if ($NoTPAData.Count -gt 0) { $NoTPAData } else {
    # Build an empty placeholder row that includes Notes and Within Timing up front,
    # then the standard column order - so headers always appear even when the sheet is empty
    $EmptyColOrder = @("Notes", "Within Timing") + $ColumnOrder
    $EmptyRow = [ordered]@{}
    foreach ($Col in $EmptyColOrder) { $EmptyRow[$Col] = "" }
    @([PSCustomObject]$EmptyRow)
}
$Pkg = $NoTPAExportData | Export-Excel -Path $OutputFilePath -WorksheetName "No TPA" -ClearSheet -PassThru
& $FormatWS $Pkg.Workbook.Worksheets["No TPA"] $BgColor $FontColor @{ "Notes" = 50 } @("Notes")
$Pkg.Save()
$Pkg.Dispose()
Write-Log "  No TPA written and formatted" -Level SUCCESS

Write-Log ""
Write-Log "========================================"
Write-Log "Processing Complete"
Write-Log "========================================"
Write-Log "Output file : $OutputFilePath" -Level SUCCESS
Write-Log "Log file    : $LogFile" -Level SUCCESS
Write-Log ""

Write-Host ""
Write-Host "Output folder: $OutputFolder" -ForegroundColor Cyan
Write-Host ""
Read-Host "Press Enter to exit"
