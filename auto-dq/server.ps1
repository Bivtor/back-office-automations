# server.ps1
# GUI-controlled pipeline server.
# Owns: HTTP listener (ports 8080 + 52731), run state, job lifecycle,
# @@STATUS parsing, Phase A PERSH-scan endpoints, defaults resolution,
# and the /assignments endpoint for the assignment userscript.
#
# Started by START.bat. No admin, no install, no system changes.
#
# STAGE FLOW:
#   Idle -> [/start called] -> PhaseA (serve RO IDs to userscript via 52731)
#        -> [all scanned or skipped] -> Assignment (Start-Job run.ps1 -JobMode)
#        -> Done / Failed / Cancelled

param(
    [int]$Port       = 8080,
    [int]$PhaseAPort = 52731,
    [string]$AppRoot = $PSScriptRoot
)

$ErrorActionPreference = "Stop"

# ==============================================================================
# DEFAULTS -- resolved once at startup, served to the GUI on /defaults
# ==============================================================================
function Get-DefaultExcelPath {
    $dl = [System.IO.Path]::Combine($env:USERPROFILE, "Downloads")
    $found = Get-ChildItem -Path $dl -Filter "xlsexport-*.xls" -ErrorAction SilentlyContinue |
             Sort-Object LastWriteTime -Descending |
             Select-Object -First 1
    if ($found) { return $found.FullName }
    return ""
}

# Next-business-day logic, mirrored from run.ps1 so the server's path preview
# and validation agree with what run.ps1 will actually name the file.
$script:HolidayFallback = @(
    "2025-01-01","2025-01-20","2025-02-17","2025-05-26","2025-06-19","2025-07-04",
    "2025-09-01","2025-10-13","2025-11-11","2025-11-27","2025-12-25","2026-01-01",
    "2026-01-19","2026-02-16","2026-05-25","2026-06-19","2026-07-03","2026-09-07",
    "2026-10-12","2026-11-11","2026-11-26","2026-12-25","2027-01-01","2027-01-18",
    "2027-02-15","2027-05-31","2027-06-19","2027-07-05","2027-09-06","2027-10-11",
    "2027-11-11","2027-11-25","2027-12-24"
)

function Get-ServerHolidays {
    param([int[]]$Years)
    $set = [System.Collections.Generic.HashSet[string]]::new()
    $ok  = $false
    foreach ($yr in $Years) {
        try {
            $resp = Invoke-RestMethod -Uri "https://date.nager.at/api/v3/PublicHolidays/$yr/US" `
                        -TimeoutSec 8 -ErrorAction Stop
            foreach ($h in $resp) { $set.Add($h.date) | Out-Null }
            $ok = $true
        } catch { $ok = $false; break }
    }
    if (-not $ok) { foreach ($d in $script:HolidayFallback) { $set.Add($d) | Out-Null } }
    return $set
}

function Get-ServerNextBusinessDay {
    $today    = (Get-Date).Date
    $holidays = Get-ServerHolidays -Years (@($today.Year, $today.Year + 1) | Select-Object -Unique)
    $d = $today.AddDays(1)
    while ($d.DayOfWeek -eq 'Saturday' -or $d.DayOfWeek -eq 'Sunday' -or
           $holidays.Contains($d.ToString("yyyy-MM-dd"))) {
        $d = $d.AddDays(1)
    }
    return $d
}

# Build the production output FOLDER from a date:
#   Q:\...\Daily Reports\{yyyy}\{MM} - {MonthName}\
# Folder and file are both keyed to the next business day, so a run on the last
# business day of a month lands in next month's folder (matching the file date).
function Get-ProdOutputFolder {
    param([datetime]$ForDate)
    $base  = "Q:\CORE\Transfers In\Reporting\Daily Queue Assignments\Daily Reports"
    $year  = $ForDate.ToString("yyyy")
    $month = "{0:D2} - {1}" -f $ForDate.Month, $ForDate.ToString("MMMM")
    return [System.IO.Path]::Combine($base, $year, $month)
}

$script:TestOutputPath = "C:\Users\USERNAME\OneDrive - COMPANY\Desktop\Projects\Auto DQ\Test_reporting_folder"

$script:DefaultExcelPath = Get-DefaultExcelPath
# Default (non-test) output is the production Q: folder for the next business day.
$script:NextBizDay        = Get-ServerNextBusinessDay
$script:DefaultOutputPath = Get-ProdOutputFolder -ForDate $script:NextBizDay

# ==============================================================================
# SHARED STATE
# ==============================================================================
$script:State = [ordered]@{
    status      = "Idle"      # Idle | Running | Done | Failed | Cancelled
    stage       = ""          # "" | PhaseA | Assignment
    phaseA      = [ordered]@{ scanned = 0; total = 0; skipped = $false }
    progress    = [ordered]@{ phase = 0; pct = 0; msg = "" }
    log         = [System.Collections.Generic.List[string]]::new()
    startedUtc  = $null
    outputFiles = [ordered]@{ workbook = ""; summary = "" }
    params      = [ordered]@{ excelPath=""; outputPath=""; exclude=""; testMode=$false }
}

# Phase A working state (separate from serialised $script:State for clarity)
$script:PhaseA = [ordered]@{
    roIds      = [System.Collections.Generic.List[string]]::new()  # ordered list to scan
    index      = 0                                                  # next to serve
    results    = [System.Collections.Generic.List[object]]::new()  # {roId, productType}
}

$script:Job    = $null
$LOG_MAX       = 500

# ==============================================================================
# HELPERS
# ==============================================================================
function Add-Log {
    param([string]$line)
    if ($null -eq $line) { return }
    $stamp = (Get-Date).ToString("HH:mm:ss")
    $script:State.log.Add("[$stamp] $line")
    while ($script:State.log.Count -gt $LOG_MAX) { $script:State.log.RemoveAt(0) }
}

function Reset-State {
    $script:State.status              = "Idle"
    $script:State.stage               = ""
    $script:State.phaseA.scanned      = 0
    $script:State.phaseA.total        = 0
    $script:State.phaseA.skipped      = $false
    $script:State.progress.phase      = 0
    $script:State.progress.pct        = 0
    $script:State.progress.msg        = ""
    $script:State.log.Clear()
    $script:State.startedUtc          = $null
    $script:State.outputFiles.workbook = ""
    $script:State.outputFiles.summary  = ""
    $script:PhaseA.roIds.Clear()
    $script:PhaseA.index   = 0
    $script:PhaseA.results.Clear()
}

function Send-Text {
    param($ctx, [string]$text, [string]$contentType = "text/plain", [int]$status = 200)
    $bytes = [System.Text.Encoding]::UTF8.GetBytes($text)
    $ctx.Response.StatusCode      = $status
    $ctx.Response.ContentType     = $contentType
    $ctx.Response.ContentLength64 = $bytes.Length
    $ctx.Response.Headers.Add("Access-Control-Allow-Origin","*")
    $ctx.Response.OutputStream.Write($bytes, 0, $bytes.Length)
    $ctx.Response.OutputStream.Close()
}

function Send-Json {
    param($ctx, $obj, [int]$status = 200)
    $json = $obj | ConvertTo-Json -Depth 8 -Compress
    Send-Text $ctx $json "application/json" $status
}

function Read-Body {
    param($req)
    if (-not $req.HasEntityBody) { return "" }
    $reader = [System.IO.StreamReader]::new($req.InputStream, $req.ContentEncoding)
    try { return $reader.ReadToEnd() } finally { $reader.Dispose() }
}

# Trim whitespace and strip a single layer of wrapping quotes (Windows
# "Copy as path" wraps in double quotes). Handles both " and '.
function Clean-Path {
    param([string]$p)
    if ($null -eq $p) { return "" }
    $p = $p.Trim()
    if ($p.Length -ge 2) {
        $first = $p[0]; $last = $p[$p.Length - 1]
        if (($first -eq '"' -and $last -eq '"') -or ($first -eq "'" -and $last -eq "'")) {
            $p = $p.Substring(1, $p.Length - 2).Trim()
        }
    }
    return $p
}

function Log-Request {
    param([string]$method, [string]$path, [string]$extra = "")
    if ($path -in @("/status","/ping")) { return }
    $t   = (Get-Date).ToString("HH:mm:ss")
    $col = switch -Regex ($path) {
        "start|result" { "Green"  }
        "cancel"       { "Yellow" }
        "shutdown"     { "Red"    }
        "phaseA|ping"  { "Cyan"   }
        default        { "Gray"   }
    }
    Write-Host ("  [{0}] {1,-5} {2}{3}" -f $t,$method,$path,$(if($extra){" -- $extra"}else{""})) -ForegroundColor $col
}

# ==============================================================================
# JOB DRAIN  (@@STATUS / @@DONE / @@FAIL marker protocol)
# ==============================================================================
function Parse-JobLine {
    param([string]$line)
    if ([string]::IsNullOrWhiteSpace($line)) { return }

    if ($line -match '^\s*@@STATUS\s+(.*)$') {
        $kv = @{}
        foreach ($m in [regex]::Matches($Matches[1],'(\w+)=((?:(?!\s+\w+=).)*)')) {
            $kv[$m.Groups[1].Value] = $m.Groups[2].Value.Trim()
        }
        if ($kv.ContainsKey('phase')) { $script:State.progress.phase = [int]$kv['phase'] }
        if ($kv.ContainsKey('pct'))   { $script:State.progress.pct   = [int]$kv['pct']   }
        if ($kv.ContainsKey('msg'))   { $script:State.progress.msg   = $kv['msg']         }
        Add-Log "[status] phase=$($script:State.progress.phase) pct=$($script:State.progress.pct) $($script:State.progress.msg)"
        return
    }
    if ($line -match '^\s*@@DONE\s*(.*)$') {
        $kv = @{}
        foreach ($m in [regex]::Matches($Matches[1],'(\w+)=((?:(?!\s+\w+=).)*)')) {
            $kv[$m.Groups[1].Value] = $m.Groups[2].Value.Trim()
        }
        if ($kv.ContainsKey('workbook')) { $script:State.outputFiles.workbook = $kv['workbook'] }
        if ($kv.ContainsKey('summary'))  { $script:State.outputFiles.summary  = $kv['summary']  }
        $script:State.progress.pct = 100
        $script:State.progress.msg = "Done"
        $script:State.status       = "Done"
        $script:State.stage        = ""
        Add-Log "Run complete."
        return
    }
    if ($line -match '^\s*@@FAIL\s*(?:msg=)?(.*)$') {
        $script:State.status = "Failed"
        $script:State.stage  = ""
        Add-Log "Run FAILED: $($Matches[1].Trim())"
        return
    }
    Add-Log $line
}

function Pump-Job {
    if ($null -eq $script:Job) { return }

    $new = Receive-Job -Job $script:Job -ErrorAction SilentlyContinue
    foreach ($item in $new) { Parse-JobLine ([string]$item) }

    if ($script:Job.State -in @('Completed','Failed','Stopped')) {
        $tail = Receive-Job -Job $script:Job -ErrorAction SilentlyContinue
        foreach ($item in $tail) { Parse-JobLine ([string]$item) }

        if ($script:State.status -eq "Running") {
            if ($script:Job.State -eq 'Completed') {
                $script:State.status = "Done"; $script:State.progress.pct = 100
                Add-Log "Job completed (no @@DONE marker; inferred)."
            } else {
                $script:State.status = "Failed"
                Add-Log "Job ended in state '$($script:Job.State)' without @@DONE."
            }
            $script:State.stage = ""
        }
        Remove-Job -Job $script:Job -Force -ErrorAction SilentlyContinue
        $script:Job = $null
    }
}

# ==============================================================================
# PHASE A HELPERS
# ==============================================================================

# Quick pre-parse: extract top-section DATA RO IDs from the export file.
# Mirrors Phase 1 of run.ps1 but only collects RO IDs -- no full row objects needed.
function Get-TopSectionRoIds {
    param([string]$filePath)
    try {
        $html     = Get-Content -Path $filePath -Raw -Encoding UTF8
        $trBlocks = [regex]::Matches($html,'(?is)<tr\b[^>]*>.*?</tr>')
        if ($trBlocks.Count -eq 0) { return @() }

        # Find header row to locate RO_ID column
        $roColIdx = -1
        foreach ($m in $trBlocks) {
            if ($m.Value -notmatch '(?i)<th\b') { continue }
            $cells = [regex]::Matches($m.Value,'(?is)<th\b[^>]*>(.*?)</th>')
            for ($i = 0; $i -lt $cells.Count; $i++) {
                $txt = ($cells[$i].Groups[1].Value -replace '(?s)<[^>]+>','').Trim()
                if ($txt -eq 'R/O ID') { $roColIdx = $i; break }
            }
            break
        }
        if ($roColIdx -lt 0) { return @() }

        $ids      = [System.Collections.Generic.List[string]]::new()
        $afterSep = $false
        foreach ($m in $trBlocks) {
            $tr = $m.Value
            if ($tr -match '(?i)<th\b') { continue }
            if ($tr -match '(?i)class\s*=\s*"[^"]*\btitle2\b') { $afterSep = $true; continue }
            if ($afterSep) { continue }
            $cells = [regex]::Matches($tr,'(?is)<td\b[^>]*>(.*?)</td>')
            if ($cells.Count -le $roColIdx) { continue }
            $roVal = ($cells[$roColIdx].Groups[1].Value -replace '(?s)<[^>]+>','') -replace '&nbsp;',' '
            $roVal = ($roVal -replace '\s+',' ').Trim()
            if ($roVal -match '^\d+$') { $ids.Add($roVal) }
        }
        return $ids.ToArray()
    } catch {
        Add-Log "Pre-parse warning: $_"
        return @()
    }
}

# Build the URL the userscript navigates to for a given RO ID
function Get-RolloverEditUrl {
    param([string]$roId)
    return "https://www.partners.example.com/SimpleAlloc/Maintenance/rolloverEdit.cfm?ID=$roId&dq=1"
}

# Write phaseA_results.json and transition to Assignment stage
function Start-AssignmentStage {
    $outputPath = $script:State.params.outputPath
    # In test mode, create the folder if missing. In production it was already
    # validated to exist in Handle-Start, so we never create on the share here.
    if ($script:State.params.testMode -and -not (Test-Path $outputPath -ErrorAction SilentlyContinue)) {
        New-Item -ItemType Directory -Path $outputPath -Force | Out-Null
    }
    $phaseAFile = Join-Path $outputPath "phaseA_results.json"
    $script:PhaseA.results | ConvertTo-Json -Depth 4 | Set-Content -Path $phaseAFile -Encoding UTF8
    Add-Log "Phase A results written: $phaseAFile ($($script:PhaseA.results.Count) entries)"
    Add-Log "Starting Assignment stage..."

    $script:State.stage = "Assignment"
    $script:State.progress.phase = 0
    $script:State.progress.pct   = 0
    $script:State.progress.msg   = "Launching pipeline..."

    $ep = $script:State.params.excelPath
    $op = $script:State.params.outputPath
    $ex = $script:State.params.exclude
    $tm = if ($script:State.params.testMode) { "true" } else { "false" }
    $runScript = Join-Path $AppRoot "run.ps1"
    $apiKey    = [string]$env:ANTHROPIC_API_KEY   # read in server's process, pass into job

    # Run run.ps1 directly in the job's runspace (NOT a nested powershell -File).
    # A nested process swallows early failures and exit codes, which is why a
    # missing module looked like a silent 4-second "completion". Here, run.ps1's
    # Write-Output lands straight in the job output stream, and any terminating
    # error is captured and surfaced as an @@FAIL the server can parse.
    $script:Job = Start-Job -ScriptBlock {
        param($runScript,$ep,$op,$ex,$tm,$apiKey)
        $ErrorActionPreference = "Stop"
        try {
            & $runScript -ExcelPath $ep -OutputPath $op -ExcludeInput $ex `
                         -TestMode $tm -ApiKey $apiKey -JobMode
        } catch {
            # Emit a marker the server's Parse-JobLine understands, then a detail line.
            Write-Output "@@FAIL msg=$($_.Exception.Message)"
            Write-Output "Stack: $($_.ScriptStackTrace -replace '[\r\n]+',' | ')"
        }
    } -ArgumentList $runScript,$ep,$op,$ex,$tm,$apiKey
}

# ==============================================================================
# REQUEST HANDLERS
# ==============================================================================

function Handle-Start {
    param($ctx)
    $bodyRaw = Read-Body $ctx.Request
    $body    = if ($bodyRaw) { $bodyRaw | ConvertFrom-Json } else { $null }

    if ($script:State.status -eq "Running") {
        Send-Json $ctx @{ ok=$false; error="A run is already in progress." } 409
        return
    }

    Reset-State

    # Resolve paths: use supplied values (quote-stripped), fall back to defaults
    $epRaw = if ($body) { Clean-Path ([string]$body.excelPath) }  else { "" }
    $opRaw = if ($body) { Clean-Path ([string]$body.outputPath) } else { "" }
    $ep = if ($epRaw) { $epRaw } else { $script:DefaultExcelPath  }
    $op = if ($opRaw) { $opRaw } else { $script:DefaultOutputPath }
    $ex = if ($body -and $body.exclude)  { [string]$body.exclude } else { "" }
    $tm = if ($body -and $body.testMode) { [bool]$body.testMode }  else { $false }

    $script:State.params.excelPath  = $ep
    $script:State.params.outputPath = $op
    $script:State.params.exclude    = $ex
    $script:State.params.testMode   = $tm
    $script:State.status            = "Running"
    $script:State.startedUtc        = (Get-Date).ToUniversalTime().ToString("o")

    Add-Log "Run started."
    Add-Log "Excel : $ep"
    Add-Log "Output: $op"
    Add-Log "Test  : $tm"

    if (-not (Test-Path $ep -ErrorAction SilentlyContinue)) {
        $script:State.status = "Failed"
        Add-Log "ERROR: export file not found: $ep"
        Send-Json $ctx @{ ok=$false; error="Export file not found: $ep" }
        return
    }

    # Output folder validation.
    #  - Production (test mode OFF): the Q: folder MUST already exist. Hard error
    #    if not, so we never silently create a stray month folder on the share.
    #  - Test mode ON: create the folder if missing (Start-AssignmentStage does this).
    if (-not $tm) {
        if (-not (Test-Path $op -ErrorAction SilentlyContinue)) {
            $script:State.status = "Failed"
            Add-Log "ERROR: output folder does not exist: $op"
            Add-Log "Create the month folder on the share first, or use Test mode."
            Send-Json $ctx @{ ok=$false; error="Output folder does not exist: $op`n`nCreate the month folder on the share first, or enable Test mode." }
            return
        }
    }

    # Pre-parse top-section RO IDs for Phase A
    $ids = Get-TopSectionRoIds -filePath $ep
    Add-Log "Pre-parse: $($ids.Count) top-section RO IDs found."

    if ($ids.Count -gt 0) {
        foreach ($id in $ids) { $script:PhaseA.roIds.Add($id) }
        $script:State.phaseA.total   = $ids.Count
        $script:State.phaseA.scanned = 0
        $script:State.stage          = "PhaseA"
        $script:State.progress.msg   = "Waiting for PERSH scan..."
        Add-Log "Stage: PhaseA. Open: $( Get-RolloverEditUrl $ids[0] )"
        Add-Log "Or click 'Skip Phase A' to proceed without PERSH detection."
    } else {
        Add-Log "No top-section rows found; skipping Phase A."
        $script:State.phaseA.skipped = $true
        Start-AssignmentStage
    }

    $firstUrl = ""
    if ($ids.Count -gt 0) { $firstUrl = Get-RolloverEditUrl $ids[0] }
    Send-Json $ctx @{ ok=$true; firstUrl=$firstUrl }
}

function Handle-Cancel {
    param($ctx)
    if ($script:State.status -ne "Running") {
        Send-Json $ctx @{ ok=$false; error="Nothing is running." }
        return
    }
    if ($script:Job) {
        Stop-Job  -Job $script:Job -ErrorAction SilentlyContinue
        Remove-Job -Job $script:Job -Force -ErrorAction SilentlyContinue
        $script:Job = $null
    }
    $script:State.status = "Cancelled"
    $script:State.stage  = ""
    Add-Log "Run cancelled by operator."
    Send-Json $ctx @{ ok=$true }
}

function Handle-SkipPhaseA {
    param($ctx)
    if ($script:State.status -ne "Running" -or $script:State.stage -ne "PhaseA") {
        Send-Json $ctx @{ ok=$false; error="Not in PhaseA stage." }
        return
    }
    Add-Log "Phase A skipped by operator. No PERSH rows will be flagged."
    $script:State.phaseA.skipped = $true
    Start-AssignmentStage
    Send-Json $ctx @{ ok=$true }
}

# ---------------------------------------------------------------------------
# Phase A userscript endpoints (port 52731, same listener via dual prefix)
# ---------------------------------------------------------------------------

# GET /ping  -- userscript polls this on each page load
function Handle-Ping {
    param($ctx)
    if ($script:State.status -ne "Running" -or $script:State.stage -ne "PhaseA") {
        Send-Json $ctx @{ ok=$false; done=$true; note="Not in PhaseA stage." }
        return
    }
    $idx   = $script:PhaseA.index
    $total = $script:PhaseA.roIds.Count
    Send-Json $ctx @{ ok=$true; done=($idx -ge $total); index=$idx; total=$total }
}

# POST /result  -- userscript posts { roId, productType } after reading the page
function Handle-Result {
    param($ctx)
    $bodyRaw = Read-Body $ctx.Request
    $body    = if ($bodyRaw) { $bodyRaw | ConvertFrom-Json } else { $null }

    if (-not $body -or -not $body.roId) {
        Send-Json $ctx @{ ok=$false; error="Missing roId" } 400
        return
    }

    $roId       = [string]$body.roId
    $productType = [string]$body.productType

    $script:PhaseA.results.Add([PSCustomObject]@{ roId=$roId; productType=$productType })
    $script:PhaseA.index++
    $script:State.phaseA.scanned = $script:PhaseA.index

    $idx   = $script:PhaseA.index
    $total = $script:PhaseA.roIds.Count

    Add-Log "PhaseA: $roId -> $productType  ($idx/$total)"

    if ($idx -ge $total) {
        # All scanned -- transition to Assignment
        Add-Log "Phase A complete."
        $script:State.phaseA.scanned = $total
        Start-AssignmentStage
        Send-Json $ctx @{ ok=$true; done=$true; index=$idx; total=$total }
    } else {
        $nextRo  = $script:PhaseA.roIds[$idx]
        $nextUrl = Get-RolloverEditUrl $nextRo
        Send-Json $ctx @{ ok=$true; done=$false; index=$idx; total=$total; nextUrl=$nextUrl }
    }
}

# GET /assignments  -- served to the future assignment userscript
function Handle-Assignments {
    param($ctx)
    # Rebuild from the last phaseA results file in the output path
    $op   = $script:State.params.outputPath
    $file = Join-Path $op "phaseA_results.json"
    if (-not (Test-Path $file -ErrorAction SilentlyContinue)) {
        Send-Json $ctx @{ ok=$false; error="No phaseA_results.json found in output path." }
        return
    }
    Send-Json $ctx @{ ok=$true; file=$file }
}

# ==============================================================================
# LISTENER SETUP (dual prefix: 8080 for GUI, 52731 for userscript)
# ==============================================================================
$listener = [System.Net.HttpListener]::new()
$listener.Prefixes.Add("http://localhost:$Port/")
$listener.Prefixes.Add("http://localhost:$PhaseAPort/")
try {
    $listener.Start()
} catch {
    Write-Host "FATAL: could not start listener." -ForegroundColor Red
    Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  Check that ports $Port and $PhaseAPort are free." -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host "  Daily Queue Server" -ForegroundColor Cyan
Write-Host "  GUI     : http://localhost:$Port/" -ForegroundColor Green
Write-Host "  Phase A : http://localhost:$PhaseAPort/ (userscript)" -ForegroundColor Green
Write-Host "  Close this window or use the GUI 'Shut down' to stop." -ForegroundColor DarkGray
Write-Host "=============================================" -ForegroundColor Cyan
Write-Host ""

$indexPath = Join-Path $AppRoot "index.html"
$running   = $true

while ($running -and $listener.IsListening) {
    $ctx    = $listener.GetContext()
    $req    = $ctx.Request
    $path   = $req.Url.AbsolutePath.TrimEnd('/')
    if ($path -eq "") { $path = "/" }
    $method = $req.HttpMethod
    $port   = $req.Url.Port

    Log-Request $method $path

    try {
        switch -Regex ("$method $path") {

            "^GET /$" {
                if (Test-Path $indexPath) {
                    Send-Text $ctx (Get-Content $indexPath -Raw -Encoding UTF8) "text/html"
                } else {
                    Send-Text $ctx "index.html not found next to server.ps1" "text/plain" 500
                }
            }

            "^GET /status$" {
                Pump-Job
                $proj = [ordered]@{
                    status      = $script:State.status
                    stage       = $script:State.stage
                    phaseA      = $script:State.phaseA
                    progress    = $script:State.progress
                    log         = @($script:State.log)
                    startedUtc  = $script:State.startedUtc
                    outputFiles = $script:State.outputFiles
                    params      = $script:State.params
                }
                Send-Json $ctx $proj
            }

            "^GET /defaults$" {
                # Refresh excel default in case a new file was downloaded since startup
                $script:DefaultExcelPath = Get-DefaultExcelPath
                $latestMtime = if ($script:DefaultExcelPath -and (Test-Path $script:DefaultExcelPath)) {
                    (Get-Item $script:DefaultExcelPath).LastWriteTime.ToString("yyyy-MM-dd")
                } else { "" }
                # Recompute next business day + prod folder (handles a server left
                # running across midnight / a date rollover).
                $script:NextBizDay        = Get-ServerNextBusinessDay
                $script:DefaultOutputPath = Get-ProdOutputFolder -ForDate $script:NextBizDay
                Send-Json $ctx @{
                    excelPath      = $script:DefaultExcelPath
                    outputPath     = $script:DefaultOutputPath       # production (toggle OFF)
                    prodOutputPath = $script:DefaultOutputPath
                    testOutputPath = $script:TestOutputPath          # test (toggle ON)
                    excelFileDate  = $latestMtime
                    today          = (Get-Date).ToString("yyyy-MM-dd")
                    nextBizDay     = $script:NextBizDay.ToString("MM-dd-yyyy")
                }
            }

            "^POST /start$"       { Handle-Start      $ctx }
            "^POST /cancel$"      { Handle-Cancel     $ctx }
            "^POST /skipPhaseA$"  { Handle-SkipPhaseA $ctx }
            "^GET  /assignments$" { Handle-Assignments $ctx }

            # Phase A userscript endpoints (callable from either port)
            "^GET /ping$"         { Handle-Ping   $ctx }
            "^POST /result$"      { Handle-Result $ctx }

            "^POST /shutdown$" {
                Add-Log "Shutdown requested."
                Send-Json $ctx @{ ok=$true }
                $running = $false
            }

            default {
                Send-Json $ctx @{ error="unknown endpoint: $method $path" } 404
            }
        }
    } catch {
        try { Send-Json $ctx @{ error="server error: $($_.Exception.Message)" } 500 } catch {}
        Write-Host "Handler error: $($_.Exception.Message)" -ForegroundColor Red
    }
}

$listener.Stop()
$listener.Close()
if ($script:Job) {
    Stop-Job  -Job $script:Job -ErrorAction SilentlyContinue
    Remove-Job -Job $script:Job -Force -ErrorAction SilentlyContinue
}
Write-Host "Server stopped." -ForegroundColor DarkGray
