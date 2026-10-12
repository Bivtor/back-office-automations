# run.ps1
#
# The export "*.xls" is an HTML table. We reproduce the ENTIRE file faithfully
# onto two styled worksheets for audit, and only COMPUTE assignments for the
# blank-Assignment rows in the TOP section (before the first separator).
#
#   "original" = raw download, every row, styled, nothing computed/deleted.
#   "modified" = same, but top-section pre-assigned rows DELETED (logged),
#                blank top-section rows get computed assignments, highlights.
#
# No Excel COM object is created anywhere.
#
# -JobMode switch: swaps Write-Host -> Write-Output, Read-Host -> @@FAIL,
#                  and emits @@STATUS / @@DONE markers for server.ps1 to drain.

param(
    [Parameter(Mandatory)][string]$ExcelPath,
    [Parameter(Mandatory)][string]$OutputPath,
    [string]$ExcludeInput    = "",
    [string]$TestMode        = "false",
    [string]$ApiKey          = "",             # accepted for compatibility (server.ps1 still passes it); no longer used
    [switch]$JobMode                           # set by server.ps1 Start-Job; no interactive I/O
)

$isTest    = ($TestMode -eq "true")
$isJobMode = $JobMode.IsPresent
$VICTOR_SHORTNAME  = "Victor"

# ---------------------------------------------------------------------------
# Logging shim: in JobMode all output goes to stdout (captured by Receive-Job).
# In console mode, Write-Host keeps colours.
# ---------------------------------------------------------------------------
function Log {
    param([string]$msg, [string]$color = "White")
    if ($isJobMode) { Write-Output $msg }
    else            { Write-Host $msg -ForegroundColor $color }
}
function Fail {
    param([string]$msg)
    if ($isJobMode) {
        # Emit the marker, then throw so the job wrapper's catch also surfaces it.
        # (Throw alone would be caught and re-emitted; emitting here guarantees the
        #  message appears even if the wrapper changes.)
        Write-Output "@@FAIL msg=$msg"
        throw $msg
    } else {
        Write-Host "  ERROR: $msg" -ForegroundColor Red
        Read-Host "  Press Enter to exit"
        exit 1
    }
}

# ==============================================================================
# PROCESSORS
# ==============================================================================
$allProcessors = @(
    [PSCustomObject]@{ shortName="Fran";    assignName="F.Foster";  selectValue="1001" },
    [PSCustomObject]@{ shortName="Victor";   assignName="V.Rinaldi"; selectValue="1002" },
    [PSCustomObject]@{ shortName="Drew Davis"; assignName="D.Davis";    selectValue="1003" },
    [PSCustomObject]@{ shortName="Evan";   assignName="E.Evans";  selectValue="1004" }
)
$PERSH_ASSIGN_NAME = "G.Grant"

$activeProcessors = $allProcessors
if (-not [string]::IsNullOrWhiteSpace($ExcludeInput)) {
    $excludeIndices = $ExcludeInput -split "," |
                      ForEach-Object { $_.Trim() } |
                      Where-Object   { $_ -match "^\d+$" } |
                      ForEach-Object { [int]$_ - 1 }
    $activeProcessors = for ($i = 0; $i -lt $allProcessors.Count; $i++) {
        if ($i -notin $excludeIndices) { $allProcessors[$i] }
    }
}

# Header label -> logical field
$HEADER_TO_FIELD = @{
    "R/O ID"="RO_ID"; "OES ID"="OES_ID"; "Client Name"="ClientName"
    "Source"="Source"; "Company"="Company"; "Received"="Received"
    "Last Modified"="LastMod"; '$ Expected'="Expected"
    "Assignment"="Assignment"; "Routed To"="RoutedTo"
}

# Styling constants
$FONT_NAME   = "Aptos Narrow"
$FONT_SIZE   = 11
$TOTAL_COLS  = 15

Log ""
Log "=============================================" Cyan
Log "  run.ps1$(if ($isJobMode){' [JobMode]'})" Cyan
if ($isTest) { Log "  *** TEST MODE ***" Yellow }
Log "  Excel  : $ExcelPath" White
Log "  Output : $OutputPath" White
Log "  Procs  : $(($activeProcessors|ForEach-Object{$_.shortName}) -join ', ')" White
Log "=============================================" Cyan
Log ""

# ==============================================================================
# PHASE 1 -- Parse the ENTIRE HTML table
# ==============================================================================
if ($isJobMode) { Write-Output "@@STATUS phase=1 pct=5 msg=Parsing export" }
Log "[1/4] Parsing export (full HTML table)..." Yellow

if (-not (Test-Path $ExcelPath)) { Fail "file not found: $ExcelPath" }

function Get-CellText {
    param([string]$inner)
    $t = $inner -replace '(?s)<[^>]+>', ''
    $t = $t -replace '&nbsp;', ' ' -replace '&amp;', '&' -replace '&lt;','<' `
            -replace '&gt;','>' -replace '&#39;',"'" -replace '&quot;','"'
    return (($t -replace '\s+', ' ').Trim())
}

function Get-RowCells {
    param([string]$trBlock)
    $tag = if ($trBlock -match '(?i)<th\b') { 'th' } else { 'td' }
    $list = [System.Collections.Generic.List[string]]::new()
    foreach ($mm in [regex]::Matches($trBlock, "(?is)<$tag\b[^>]*>(.*?)</$tag>")) {
        $list.Add( (Get-CellText $mm.Groups[1].Value) )
    }
    return ,$list.ToArray()
}

try {
    $html     = Get-Content -Path $ExcelPath -Raw -Encoding UTF8
    $trBlocks = [regex]::Matches($html, '(?is)<tr\b[^>]*>.*?</tr>')
    if ($trBlocks.Count -eq 0) { throw "No <tr> rows found -- export format changed." }

    # Resolve header -> column map
    $colMap = @{}
    foreach ($m in $trBlocks) {
        if ($m.Value -notmatch '(?i)<th\b') { continue }
        $hdr = Get-RowCells $m.Value
        for ($i = 0; $i -lt $hdr.Count; $i++) {
            if ($HEADER_TO_FIELD.ContainsKey($hdr[$i])) {
                $f = $HEADER_TO_FIELD[$hdr[$i]]
                if (-not $colMap.ContainsKey($f)) { $colMap[$f] = $i }
            }
        }
        break
    }
    $missing = @()
    foreach ($f in $HEADER_TO_FIELD.Values) { if (-not $colMap.ContainsKey($f)) { $missing += $f } }
    if ($missing.Count -gt 0) { throw "Header missing column(s): $($missing -join ', '). Format changed." }

    Log "  Columns OK (Company=$($colMap['Company']) Assignment=$($colMap['Assignment']))" DarkGray

    # Column type map, keyed by 0-indexed column position. Built once from
    # colMap so both the "original" and "modified" sheets apply identical
    # typed values -- no duplicated per-sheet logic.
    #   "int"  -> whole number (RO/OES IDs)
    #   "num"  -> currency/number ($ Expected)
    #   "date" -> date (Received, Last Modified)
    #   "text" -> everything else (ClientName, Source, Assignment, RoutedTo, Company)
    $script:ColumnTypes = @{}
    foreach ($f in $HEADER_TO_FIELD.Values) {
        if (-not $colMap.ContainsKey($f)) { continue }
        $idx = $colMap[$f]
        $script:ColumnTypes[$idx] = switch ($f) {
            "RO_ID"      { "int"  }
            "OES_ID"     { "int"  }
            "Expected"   { "num"  }
            "Received"   { "date" }
            "LastMod"    { "date" }
            default      { "text" }
        }
    }

    # Tag every <tr>
    $rows     = [System.Collections.Generic.List[object]]::new()
    $afterSep = $false
    $roIdx    = $colMap['RO_ID']
    $cmpIdx   = $colMap['Company']
    $asgIdx   = $colMap['Assignment']
    $cliIdx   = $colMap['ClientName']

    foreach ($m in $trBlocks) {
        $tr = $m.Value
        if ($tr -match '(?i)<th\b') {
            $rows.Add([PSCustomObject]@{ Type="HEADER"; Cells=(Get-RowCells $tr); Label=$null; AfterSep=$afterSep })
            continue
        }
        if ($tr -match '(?i)class\s*=\s*"[^"]*\btitle2\b') {
            $lbl = (Get-RowCells $tr)[0]
            $rows.Add([PSCustomObject]@{ Type="SEPARATOR"; Cells=@($lbl); Label=$lbl; AfterSep=$afterSep })
            $afterSep = $true
            continue
        }
        $cells = Get-RowCells $tr
        if ($cells.Count -eq 0) { continue }
        $roVal = if ($roIdx -lt $cells.Count) { [string]$cells[$roIdx] } else { "" }
        if ($roVal -match '^\d+$') {
            $rows.Add([PSCustomObject]@{ Type="DATA"; Cells=$cells; Label=$null; AfterSep=$afterSep })
        } else {
            $rows.Add([PSCustomObject]@{ Type="DASH"; Cells=$cells; Label=$null; AfterSep=$afterSep })
        }
    }

    $dataTop = @($rows | Where-Object { $_.Type -eq "DATA" -and -not $_.AfterSep })
    $dataAll = @($rows | Where-Object { $_.Type -eq "DATA" })
    Log "  Total rows: $($rows.Count)  | DATA: $($dataAll.Count)  | top-section DATA: $($dataTop.Count)" White
    if ($dataAll.Count -eq 0) { throw "No numeric DATA rows parsed." }

} catch {
    Fail "parsing file: $_"
}

# ==============================================================================
# PHASE 2 -- Classify TOP-SECTION rows
# ==============================================================================
if ($isJobMode) { Write-Output "@@STATUS phase=2 pct=25 msg=Classifying top section" }
Log ""
Log "[2/4] Classifying top section..." Yellow

$phaseAFile = Join-Path $OutputPath "phaseA_results.json"
$phaseAMap  = @{}
if (Test-Path $phaseAFile) {
    $phaseARaw = Get-Content $phaseAFile -Raw | ConvertFrom-Json
    foreach ($e in $phaseARaw) {
        $phaseAMap[[string]$e.roId] = [string]$e.productType
    }
    Log "  Phase A loaded: $($phaseAMap.Count) scanned" Green
} else {
    Log "  WARNING: phaseA_results.json missing -- PERSH not flagged." Yellow
}

$toCompute   = [System.Collections.Generic.List[object]]::new()
$internalTop = [System.Collections.Generic.List[string]]::new()
$pershTop    = [System.Collections.Generic.List[string]]::new()
$deletedTop  = [System.Collections.Generic.List[string]]::new()

foreach ($r in $dataTop) {
    $ro  = [string]$r.Cells[$roIdx]
    $cmp = [string]$r.Cells[$cmpIdx]
    $asg = [string]$r.Cells[$asgIdx]

    if (-not [string]::IsNullOrWhiteSpace($asg)) { $deletedTop.Add($ro); continue }
    if ($cmp -match "(?i)internal\s*company")  { $internalTop.Add($ro); continue }
    if ($phaseAMap.ContainsKey($ro) -and $phaseAMap[$ro] -match "(?i)persh") { $pershTop.Add($ro); continue }
    $toCompute.Add($r)
}

Log "  To compute (blank)        : $($toCompute.Count)" Green
Log "  Internal (in-house)       : $($internalTop.Count)" Cyan
Log "  PERSH                     : $($pershTop.Count)" Magenta
Log "  Deleted (pre-assigned)    : $($deletedTop.Count)" Yellow
if ($deletedTop.Count -gt 0) {
    Log ("    Deleted RO IDs: {0}" -f ($deletedTop -join ", ")) DarkYellow
}

# Holiday / next-business-day
$HOLIDAY_API     = "https://date.nager.at/api/v3/PublicHolidays/{year}/US"
$HOLIDAY_TIMEOUT = 8
$HOLIDAY_FALLBACK = @(
    "2025-01-01","2025-01-20","2025-02-17","2025-05-26","2025-06-19","2025-07-04",
    "2025-09-01","2025-10-13","2025-11-11","2025-11-27","2025-12-25","2026-01-01",
    "2026-01-19","2026-02-16","2026-05-25","2026-06-19","2026-07-03","2026-09-07",
    "2026-10-12","2026-11-11","2026-11-26","2026-12-25","2027-01-01","2027-01-18",
    "2027-02-15","2027-05-31","2027-06-19","2027-07-05","2027-09-06","2027-10-11",
    "2027-11-11","2027-11-25","2027-12-24"
)

function Get-USHolidays {
    param([int[]]$Years)
    $set = [System.Collections.Generic.HashSet[string]]::new()
    $ok  = $false
    foreach ($yr in $Years) {
        try {
            $resp = Invoke-RestMethod -Uri ($HOLIDAY_API -replace '\{year\}',$yr) `
                        -TimeoutSec $HOLIDAY_TIMEOUT -ErrorAction Stop
            foreach ($h in $resp) { $set.Add($h.date) | Out-Null }
            $ok = $true
        } catch { $ok = $false; break }
    }
    if (-not $ok) { foreach ($d in $HOLIDAY_FALLBACK) { $set.Add($d) | Out-Null } }
    return $set
}

function Get-NextBusinessDay {
    param([datetime]$From, [System.Collections.Generic.HashSet[string]]$Holidays)
    $d = $From.AddDays(1)
    while ($d.DayOfWeek -eq 'Saturday' -or $d.DayOfWeek -eq 'Sunday' -or
           $Holidays.Contains($d.ToString("yyyy-MM-dd"))) {
        $d = $d.AddDays(1)
    }
    return $d
}

$today      = (Get-Date).Date
$holidays   = Get-USHolidays -Years (@($today.Year, $today.Year + 1) | Select-Object -Unique)
$nextBizDay = Get-NextBusinessDay -From $today -Holidays $holidays
Log "  Next business day: $($nextBizDay.ToString('MM/dd/yyyy'))" DarkGray

# ==============================================================================
# PHASE 3 -- Deterministic OES-grouped sequential assignment
# ==============================================================================
# Mirrors the manual process:
#   * Only the to-compute rows (blank Assignment, NOT internal, NOT PERSH) are
#     distributed. Internal stays blank/yellow; PERSH goes to G.Grant.
#   * Work is grouped by OES ID -- one OES (all its rollovers) goes to ONE
#     processor, never split. The COUNT that matters is rollover ROWS.
#   * Targets are an even split of total rollover rows across active processors.
#     An indivisible total spreads the remainder ONE extra each: Victor takes
#     the first +1 (guaranteeing him a full share), then the rest of the +1s go
#     down the fixed order. Net effect: one person ends up short, never Victor
#     ending up with everything piled on.
#       39 / 4 -> Victor 10, Fran 10, Drew 10, Evan 9   (old rule: Victor 12)
#       25 / 4 -> Victor 7,  Fran 6,  Drew 6,  Evan 6
#   * Fill goes down the line: for each processor we add OES groups until adding
#     the next group would EXCEED their target. If the next group is too big, we
#     look ahead for a smaller group that lands them exactly on target and take
#     that instead, leaving the big group for a later processor. Any groups that
#     can't be placed cleanly go to whoever has the FEWEST rows at that point
#     (ties prefer non-Victor, then fixed order).
if ($isJobMode) { Write-Output "@@STATUS phase=3 pct=50 msg=Generating assignments" }
Log ""
Log "[3/4] Generating assignments..." Yellow

$procNames    = @($activeProcessors | ForEach-Object { $_.shortName })
$victorActive = $procNames -contains $VICTOR_SHORTNAME

# Final map: processor shortName -> List[RowId]
$assignMap = @{}
foreach ($p in $procNames) { $assignMap[$p] = [System.Collections.Generic.List[string]]::new() }

if ($toCompute.Count -gt 0 -and $procNames.Count -gt 0) {

    # --- Group to-compute rows by OES ID, preserving first-appearance order ---
    # A blank OES ID can't be merged with other blanks (unrelated), so each blank
    # becomes its own singleton group keyed by RowId.
    $oesIdx     = $colMap['OES_ID']
    $groupOrder = [System.Collections.Generic.List[string]]::new()  # group keys in order seen
    $groupRows  = @{}                                               # key -> List[RowId]

    foreach ($r in $toCompute) {
        $ro  = [string]$r.Cells[$roIdx]
        $oes = if ($r.Cells.Count -gt $oesIdx) { ([string]$r.Cells[$oesIdx]).Trim() } else { "" }
        $key = if ([string]::IsNullOrWhiteSpace($oes)) { "RO:$ro" } else { "OES:$oes" }
        if (-not $groupRows.ContainsKey($key)) {
            $groupRows[$key] = [System.Collections.Generic.List[string]]::new()
            $groupOrder.Add($key)
        }
        $groupRows[$key].Add($ro)
    }

    $totalRows = $toCompute.Count
    $nProc     = $procNames.Count

    # --- Targets: even split, remainder spread one extra each ---
    # Victor gets the first +1 so he is never the one left short; remaining
    # +1s go down the fixed order to the other processors.
    $base      = [int][math]::Floor($totalRows / $nProc)
    $remainder = $totalRows - ($base * $nProc)
    $target    = @{}
    foreach ($p in $procNames) { $target[$p] = $base }
    $plusOrder = @()
    if ($victorActive) { $plusOrder += $VICTOR_SHORTNAME }
    $plusOrder += @($procNames | Where-Object { $_ -ne $VICTOR_SHORTNAME })
    for ($i = 0; $i -lt $remainder; $i++) { $target[$plusOrder[$i]] += 1 }

    Log ("  Targets: {0}" -f (($procNames | ForEach-Object { "$_=$($target[$_])" }) -join ", ")) DarkGray

    # --- Greedy fill, down the line, with exact-fit lookahead ---
    $assignedKeys = @{}   # group key -> $true once placed
    $sizeOf       = { param($k) $groupRows[$k].Count }

    foreach ($p in $procNames) {
        # Add groups until this processor reaches their target.
        while ($assignMap[$p].Count -lt $target[$p]) {
            $room = $target[$p] - $assignMap[$p].Count

            # First unplaced group in order that FITS within remaining room.
            $pick = $null
            foreach ($k in $groupOrder) {
                if ($assignedKeys.ContainsKey($k)) { continue }
                if ($groupRows[$k].Count -le $room) { $pick = $k; break }
            }

            if ($null -eq $pick) {
                # No remaining group fits without overflowing this processor.
                # Leave them short; remaining groups go to later processors (or
                # ultimately the extras target).
                break
            }

            $assignedKeys[$pick] = $true
            foreach ($rid in $groupRows[$pick]) { $assignMap[$p].Add($rid) }
        }
    }

    # --- Any groups still unplaced (didn't fit anyone cleanly) ---
    # Send each to whoever currently holds the FEWEST rows; ties prefer
    # non-Victor, then fixed order. Victor never absorbs overflow by default.
    foreach ($k in $groupOrder) {
        if ($assignedKeys.ContainsKey($k)) { continue }
        $assignedKeys[$k] = $true
        $dest = $null
        foreach ($p in $procNames) {
            if ($null -eq $dest) { $dest = $p; continue }
            $pc = $assignMap[$p].Count
            $dc = $assignMap[$dest].Count
            if ($pc -lt $dc) { $dest = $p; continue }
            if ($pc -eq $dc -and $dest -eq $VICTOR_SHORTNAME -and $p -ne $VICTOR_SHORTNAME) { $dest = $p }
        }
        foreach ($rid in $groupRows[$k]) { $assignMap[$dest].Add($rid) }
    }
}

# RowId -> assignName lookup
$assignByRO = @{}
if ($assignMap) {
    foreach ($p in $activeProcessors) {
        if ($assignMap.ContainsKey($p.shortName)) {
            foreach ($rid in $assignMap[$p.shortName]) { $assignByRO[[string]$rid] = $p.assignName }
        }
    }
}
foreach ($ro in $pershTop) { $assignByRO[$ro] = $PERSH_ASSIGN_NAME }

# ==============================================================================
# PHASE 4 -- Write the two styled sheets
# ==============================================================================
if ($isJobMode) { Write-Output "@@STATUS phase=4 pct=75 msg=Writing workbook" }
Log ""
Log "[4/4] Writing styled workbook..." Yellow

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Fail "ImportExcel not installed. Run: Install-Module ImportExcel -Scope CurrentUser"
}
Import-Module ImportExcel -ErrorAction Stop
if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath -Force | Out-Null }

$outFile = Join-Path $OutputPath ("OES Queue Assignment {0}.xlsx" -f $nextBizDay.ToString("MM-dd-yyyy"))
if (Test-Path $outFile) { Remove-Item $outFile -Force }

$headerRowObj = $rows | Where-Object { $_.Type -eq "HEADER" } | Select-Object -First 1
$headerCells  = @($headerRowObj.Cells)

# Set a single cell's value. PowerShell's overload resolution on EPPlus's
# two-index indexer ($ws.Cells[$row,$col]) can fail with "property 'Value'
# cannot be found" depending on the OfficeOpenXml assembly version. The string
# A1-address indexer ($ws.Cells["A1"]) is unambiguous and resolves on every
# EPPlus version -- it's the same form the font-styling line uses successfully.
function ConvertTo-ColumnLetter {
    param([int]$col)
    $s = ""
    while ($col -gt 0) {
        $rem = ($col - 1) % 26
        $s   = [char](65 + $rem) + $s
        $col = [int][math]::Floor(($col - 1) / 26)
    }
    return $s
}
function Set-CellValue {
    param($ws, [int]$row, [int]$col, $value)
    $addr = (ConvertTo-ColumnLetter $col) + $row
    $ws.Cells[$addr].Value = $value
}

# Type-aware cell setter for DATA rows. Converts the raw string from the export
# into a real typed value (int / double / DateTime) per the column's type from
# $script:ColumnTypes, so Excel stores genuine numbers/dates instead of text --
# this is what eliminates the "number stored as text" warnings. Used identically
# for both the "original" and "modified" sheets (single source of truth). Any
# parse failure falls back to the original text rather than throwing, since a
# handful of malformed/blank cells shouldn't kill the whole run.
function Set-TypedCellValue {
    param($ws, [int]$row, [int]$colIdx0, [string]$rawValue)

    $col     = $colIdx0 + 1
    $addr    = (ConvertTo-ColumnLetter $col) + $row
    $type    = if ($script:ColumnTypes.ContainsKey($colIdx0)) { $script:ColumnTypes[$colIdx0] } else { "text" }
    $trimmed = if ($null -eq $rawValue) { "" } else { $rawValue.Trim() }

    if ([string]::IsNullOrWhiteSpace($trimmed)) {
        $ws.Cells[$addr].Value = ""
        return
    }

    switch ($type) {
        "int" {
            $clean  = $trimmed -replace '[^\d\-]', ''
            $parsed = 0
            if ($clean -ne "" -and [int]::TryParse($clean, [ref]$parsed)) {
                $ws.Cells[$addr].Value = $parsed
            } else {
                $ws.Cells[$addr].Value = $trimmed   # fallback: leave as text
            }
        }
        "num" {
            # Strip $ , and whitespace so "$1,234.50" parses as 1234.50
            $clean  = $trimmed -replace '[\$,\s]', ''
            $parsed = 0.0
            if ($clean -ne "" -and [double]::TryParse($clean, [ref]$parsed)) {
                $ws.Cells[$addr].Value = $parsed
                $ws.Cells[$addr].Style.Numberformat.Format = '$#,##0.00'
            } else {
                $ws.Cells[$addr].Value = $trimmed
            }
        }
        "date" {
            $parsed = [datetime]::MinValue
            if ([datetime]::TryParse($trimmed, [ref]$parsed)) {
                $ws.Cells[$addr].Value = $parsed
                $ws.Cells[$addr].Style.Numberformat.Format = 'MM/dd/yyyy'
            } else {
                $ws.Cells[$addr].Value = $trimmed   # fallback: leave as text
            }
        }
        default {
            $ws.Cells[$addr].Value = $trimmed
        }
    }
}

# Full-width range for one row, e.g. "A5:O5", via the same unambiguous string
# indexer. Used for header/separator/highlight styling.
# NOTE: the leading comma (unary array operator) stops PowerShell from unrolling
# the multi-cell ExcelRange into an array of individual cells on return -- without
# it, the caller gets an object[] and ".Style.Font.Bold" fails to bind.
function Get-RowRange {
    param($ws, [int]$row, [int]$cols)
    $addr = "A$row`:" + (ConvertTo-ColumnLetter $cols) + $row
    return ,$ws.Cells[$addr]
}

function Write-SheetRows {
    param($ws, [bool]$isModified)

    $excelRow = 1
    for ($c = 0; $c -lt $TOTAL_COLS; $c++) {
        Set-CellValue $ws $excelRow ($c + 1) $headerCells[$c]
    }
    $hr = Get-RowRange $ws $excelRow $TOTAL_COLS
    $hr.Style.Font.Bold = $true
    $hr.Style.Font.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml("#FFFFFF"))
    $hr.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $hr.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#6699CC"))
    $excelRow++

    foreach ($r in $rows) {
        if ($r.Type -eq "HEADER") { continue }

        if ($r.Type -eq "SEPARATOR") {
            Set-CellValue $ws $excelRow 1 $r.Label
            $sp = Get-RowRange $ws $excelRow $TOTAL_COLS
            $sp.Merge = $true
            $sp.Style.Font.Bold = $true
            $sp.Style.Font.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml("#FFFFFF"))
            $sp.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
            $sp.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#006699"))
            $excelRow++
            continue
        }

        if ($r.Type -eq "DASH") {
            # No R/O ID attached (dash or blank). Keep on "original" (full
            # fidelity to the source export); drop from "modified".
            if ($isModified) { continue }
        }

        $ro  = if ($r.Cells.Count -gt $roIdx) { [string]$r.Cells[$roIdx] } else { "" }
        $asg = if ($r.Cells.Count -gt $asgIdx) { [string]$r.Cells[$asgIdx] } else { "" }

        # "modified" drops a DATA row that already has an Assignment ONLY in
        # the top section. After-separator rows are left unaltered either way.
        if ($isModified -and $r.Type -eq "DATA" -and -not $r.AfterSep -and -not [string]::IsNullOrWhiteSpace($asg)) {
            continue
        }

        for ($c = 0; $c -lt $TOTAL_COLS; $c++) {
            $val = if ($c -lt $r.Cells.Count) { [string]$r.Cells[$c] } else { "" }
            Set-TypedCellValue $ws $excelRow $c $val
        }

        if ($isModified -and $r.Type -eq "DATA" -and -not $r.AfterSep) {
            if ($assignByRO.ContainsKey($ro)) {
                Set-TypedCellValue $ws $excelRow $asgIdx $assignByRO[$ro]
            }
            if ($internalTop -contains $ro) {
                $rng = Get-RowRange $ws $excelRow $TOTAL_COLS
                $rng.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $rng.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#FFFF00"))
            } elseif ($pershTop -contains $ro) {
                $rng = Get-RowRange $ws $excelRow $TOTAL_COLS
                $rng.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $rng.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#F2CEEF"))
            }
        }
        $excelRow++
    }

    $used = $ws.Dimension
    if ($used) {
        $all = $ws.Cells[$used.Address]
        $all.Style.Font.Name = $FONT_NAME
        $all.Style.Font.Size = $FONT_SIZE

        # Autosize everything first, then pin the blank spacer columns E-I
        # (1-indexed cols 5-9) to a small fixed width -- autofit would otherwise
        # shrink them to near-zero since they're empty, but we want them
        # narrow/visible rather than auto-sized to the header text.
        $ws.Cells[$used.Address].AutoFitColumns()
        foreach ($col in 5..9) {
            $ws.Column($col).Width = 5
        }
    }
}

$pkg = New-Object OfficeOpenXml.ExcelPackage
$wsM = $pkg.Workbook.Worksheets.Add("modified")
$wsO = $pkg.Workbook.Worksheets.Add("original")
Write-SheetRows -ws $wsM -isModified $true
Write-SheetRows -ws $wsO -isModified $false
$pkg.SaveAs([System.IO.FileInfo]::new($outFile))
$pkg.Dispose()

Log "  + Saved: $outFile" Green

# Audit CSV -- only in test mode. In production it just gets deleted each day,
# so we skip writing it entirely.
$summaryPath = ""
if ($isTest) {
    $summary = [System.Collections.Generic.List[object]]::new()
    foreach ($r in $toCompute) {
        $ro = [string]$r.Cells[$roIdx]
        $summary.Add([PSCustomObject]@{
            RO_ID      = $ro
            OES_ID     = $r.Cells[$colMap['OES_ID']]
            ClientName = $r.Cells[$cliIdx]
            Company    = $r.Cells[$cmpIdx]
            Expected   = $r.Cells[$colMap['Expected']]
            Assignment = ($assignByRO[$ro])
            Action     = "COMPUTED"
        })
    }
    foreach ($ro in $pershTop)    { $summary.Add([PSCustomObject]@{ RO_ID=$ro;OES_ID="";ClientName="";Company="";Expected="";Assignment=$PERSH_ASSIGN_NAME;Action="PERSH" }) }
    foreach ($ro in $internalTop) { $summary.Add([PSCustomObject]@{ RO_ID=$ro;OES_ID="";ClientName="";Company="";Expected="";Assignment="";Action="INTERNAL (yellow, blank)" }) }
    foreach ($ro in $deletedTop)  { $summary.Add([PSCustomObject]@{ RO_ID=$ro;OES_ID="";ClientName="";Company="";Expected="";Assignment="";Action="DELETED from modified (pre-assigned)" }) }

    $summaryPath = Join-Path $OutputPath ("DQ_Summary_{0}.csv" -f $nextBizDay.ToString("MM-dd-yyyy"))
    $summary | Export-Csv -Path $summaryPath -NoTypeInformation -Encoding UTF8
}

Log ""
Log "=============================================" Cyan
Log "  DAILY QUEUE SUMMARY" Cyan
Log "=============================================" Cyan
Log ("  Computed (to processors) : {0}" -f $toCompute.Count) Green
Log ("  PERSH (G.Grant)         : {0}" -f $pershTop.Count) Magenta
Log ("  Internal (yellow, blank) : {0}" -f $internalTop.Count) Yellow
Log ("  Deleted from modified    : {0}" -f $deletedTop.Count) DarkYellow
Log ""
if ($assignMap) {
    Log "  Assignments per processor:" White
    foreach ($p in $activeProcessors) {
        $n  = if ($assignMap.ContainsKey($p.shortName)) { $assignMap[$p.shortName].Count } else { 0 }
        Log ("    {0,-10} {1,3}" -f $p.shortName, $n) White
    }
}
Log ""
Log "  Workbook : $outFile" Green
if ($summaryPath) { Log "  Summary  : $summaryPath" Green }
Log ""

# Clean up the Phase A scratch file now that the run succeeded. Both output
# files are written at this point, so the JSON has served its purpose.
$phaseACleanup = Join-Path $OutputPath "phaseA_results.json"
if (Test-Path $phaseACleanup) {
    try {
        Remove-Item $phaseACleanup -Force
        Log "  Cleaned up phaseA_results.json" DarkGray
    } catch {
        Log "  NOTE: could not remove phaseA_results.json: $_" Yellow
    }
}

if ($isJobMode) {
    $doneExtra = "computed=$($toCompute.Count) forDate=$($nextBizDay.ToString('MM/dd/yyyy'))"
    if ($summaryPath) {
        Write-Output "@@DONE workbook=$outFile summary=$summaryPath $doneExtra"
    } else {
        Write-Output "@@DONE workbook=$outFile $doneExtra"
    }
} else {
    Read-Host "  Press Enter to close"
}
