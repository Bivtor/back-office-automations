# server.ps1
# GUI-controlled Outstanding Rollover Report server.
#
# Owns: the HTTP listener (port 8080), run state, the run.ps1 job lifecycle,
# @@STATUS / @@COUNT / @@PERSH / @@NEARMISS / @@DONE / @@FAIL parsing, and
# default path resolution (the guessing that generate_command.ps1 used to do
# at the console prompt).
#
# Started by START.bat. No admin, no install, no system changes.
#
# FLOW:
#   Idle -> [POST /start] -> Running (Start-Job run.ps1 -JobMode)
#        -> Done / Failed / Cancelled
#
# The browser polls GET /status once a second. Every poll drains whatever the
# job has written since the last one, so the log and progress bar are live
# without any push/websocket machinery.

param(
    [int]$Port       = 8080,
    [string]$AppRoot = $PSScriptRoot
)

$ErrorActionPreference = "Stop"

# ==============================================================================
# DEFAULTS -- the path guessing lifted from generate_command.ps1.
# Resolved on demand so a file downloaded after startup is still picked up.
# ==============================================================================

function Get-DownloadsFolder {
    $profileDir = $env:USERPROFILE
    if ([string]::IsNullOrWhiteSpace($profileDir)) {
        $profileDir = "C:\Users\$env:USERNAME"
    }
    return [System.IO.Path]::Combine($profileDir, "Downloads")
}

# Raw input: the exact filename first (that is what the export always lands as),
# then the newest file whose name looks like a custom rollover search.
function Get-DefaultRawInput {
    $dl = Get-DownloadsFolder
    if (-not (Test-Path $dl -ErrorAction SilentlyContinue)) { return "" }

    $exact = [System.IO.Path]::Combine($dl, "Custom Rollover Search Results.xlsx")
    if (Test-Path $exact -ErrorAction SilentlyContinue) { return $exact }

    $found = Get-ChildItem -Path $dl -File -ErrorAction SilentlyContinue |
             Where-Object { $_.Name -match '(?i)custom.*rollover.*search' -and $_.Extension -match '(?i)\.xlsx?$' } |
             Sort-Object LastWriteTime -Descending |
             Select-Object -First 1
    if ($found) { return $found.FullName }
    return ""
}

# Old report: newest file matching "outstanding rollover" -- same regex
# generate_command.ps1 used.
function Get-DefaultOldReport {
    $dl = Get-DownloadsFolder
    if (-not (Test-Path $dl -ErrorAction SilentlyContinue)) { return "" }

    $found = Get-ChildItem -Path $dl -File -ErrorAction SilentlyContinue |
             Where-Object { $_.Name -match '(?i)outstanding\s+rollover' -and $_.Extension -match '(?i)\.xlsx?$' } |
             Sort-Object LastWriteTime -Descending |
             Select-Object -First 1
    if ($found) { return $found.FullName }
    return ""
}

# Recent spreadsheets in Downloads, newest first. The GUI hangs these off a
# <datalist> on both path fields so picking a different file is a dropdown
# rather than a paste. Browsers cannot hand us a real full path from a file
# input, so this is the closest thing to a file picker we can offer.
function Get-RecentSpreadsheets {
    $dl = Get-DownloadsFolder
    if (-not (Test-Path $dl -ErrorAction SilentlyContinue)) { return @() }
    return @(
        Get-ChildItem -Path $dl -File -ErrorAction SilentlyContinue |
        Where-Object { $_.Extension -match '(?i)\.xlsx?$' -and $_.Name -notlike '~$*' } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 12 |
        ForEach-Object {
            [ordered]@{
                path     = $_.FullName
                name     = $_.Name
                modified = $_.LastWriteTime.ToString("MM/dd HH:mm")
            }
        }
    )
}

function Get-FileDateString {
    param([string]$path)
    if ($path -and (Test-Path $path -ErrorAction SilentlyContinue)) {
        return (Get-Item $path).LastWriteTime.ToString("yyyy-MM-dd")
    }
    return ""
}

# ==============================================================================
# SHARED STATE
# ==============================================================================
$script:State = [ordered]@{
    status      = "Idle"     # Idle | Running | Done | Failed | Cancelled
    progress    = [ordered]@{ phase = 0; pct = 0; msg = "" }
    log         = [System.Collections.Generic.List[string]]::new()
    counts      = [System.Collections.Generic.List[object]]::new()   # {label, rows}
    persh       = [System.Collections.Generic.List[string]]::new()   # manual NetX check
    nearMiss    = [System.Collections.Generic.List[string]]::new()   # review candidates
    startedUtc  = $null
    outputFiles = [ordered]@{ workbook = ""; folder = ""; log = "" }
    params      = [ordered]@{ rawInput = ""; oldReport = "" }
}

$script:Job = $null
$LOG_MAX    = 500
$PHASE_MAX  = 6

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
    $script:State.status               = "Idle"
    $script:State.progress.phase       = 0
    $script:State.progress.pct         = 0
    $script:State.progress.msg         = ""
    $script:State.log.Clear()
    $script:State.counts.Clear()
    $script:State.persh.Clear()
    $script:State.nearMiss.Clear()
    $script:State.startedUtc           = $null
    $script:State.outputFiles.workbook = ""
    $script:State.outputFiles.folder   = ""
    $script:State.outputFiles.log      = ""
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

# Trim whitespace and strip one layer of wrapping quotes. Windows "Copy as
# path" wraps in double quotes, and generate_command.ps1 stripped both quote
# characters off pasted input -- same behaviour here.
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
    param([string]$method, [string]$path)
    if ($path -in @("/status")) { return }
    $t   = (Get-Date).ToString("HH:mm:ss")
    $col = switch -Regex ($path) {
        "start"    { "Green"  }
        "cancel"   { "Yellow" }
        "shutdown" { "Red"    }
        "open"     { "Cyan"   }
        default    { "Gray"   }
    }
    Write-Host ("  [{0}] {1,-5} {2}" -f $t, $method, $path) -ForegroundColor $col
}

# ==============================================================================
# JOB DRAIN  (marker protocol -- see the header of run.ps1)
# ==============================================================================
function Get-MarkerPairs {
    # Splits "key=value key2=value with spaces" into a hashtable. The lookahead
    # lets a value contain spaces (file paths do) while still terminating at the
    # next "  word=" pair.
    param([string]$text)
    $kv = @{}
    foreach ($m in [regex]::Matches($text, '(\w+)=((?:(?!\s+\w+=).)*)')) {
        $kv[$m.Groups[1].Value] = $m.Groups[2].Value.Trim()
    }
    return $kv
}

function Parse-JobLine {
    param([string]$line)
    if ([string]::IsNullOrWhiteSpace($line)) { return }

    if ($line -match '^\s*@@STATUS\s+(.*)$') {
        $kv = Get-MarkerPairs $Matches[1]
        if ($kv.ContainsKey('phase')) { $script:State.progress.phase = [int]$kv['phase'] }
        if ($kv.ContainsKey('pct'))   { $script:State.progress.pct   = [int]$kv['pct']   }
        if ($kv.ContainsKey('msg'))   { $script:State.progress.msg   = $kv['msg']        }
        Add-Log "[$($script:State.progress.pct)%] $($script:State.progress.msg)"
        return
    }

    if ($line -match '^\s*@@COUNT\s+(.*)$') {
        $payload = $Matches[1]
        $split   = $payload.LastIndexOf('|')
        if ($split -gt 0) {
            $label = $payload.Substring(0, $split).Trim()
            $rows  = 0
            [void][int]::TryParse($payload.Substring($split + 1).Trim(), [ref]$rows)
            $script:State.counts.Add([PSCustomObject]@{ label = $label; rows = $rows })
        }
        return
    }

    if ($line -match '^\s*@@PERSH\s+(.*)$') {
        $script:State.persh.Add($Matches[1].Trim())
        return
    }

    if ($line -match '^\s*@@NEARMISS\s+(.*)$') {
        $script:State.nearMiss.Add($Matches[1].Trim())
        return
    }

    if ($line -match '^\s*@@DONE\s*(.*)$') {
        $kv = Get-MarkerPairs $Matches[1]
        if ($kv.ContainsKey('workbook')) { $script:State.outputFiles.workbook = $kv['workbook'] }
        if ($kv.ContainsKey('folder'))   { $script:State.outputFiles.folder   = $kv['folder']   }
        if ($kv.ContainsKey('log'))      { $script:State.outputFiles.log      = $kv['log']      }
        $script:State.progress.phase = $PHASE_MAX
        $script:State.progress.pct   = 100
        $script:State.progress.msg   = "Done"
        $script:State.status         = "Done"
        Add-Log "Report generation complete."
        return
    }

    if ($line -match '^\s*@@FAIL\s*(?:msg=)?(.*)$') {
        $script:State.status = "Failed"
        Add-Log "RUN FAILED: $($Matches[1].Trim())"
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
                # run.ps1 finished without emitting @@DONE. That means it hit a
                # `return` in the config guards, or exited early. Treat it as a
                # failure rather than reporting a success with no workbook.
                $script:State.status = "Failed"
                Add-Log "Job ended without a completion marker. Check the console window and the log file."
            } else {
                $script:State.status = "Failed"
                Add-Log "Job ended in state '$($script:Job.State)' without completing."
            }
        }
        Remove-Job -Job $script:Job -Force -ErrorAction SilentlyContinue
        $script:Job = $null
    }
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

    $rawIn = if ($body) { Clean-Path ([string]$body.rawInput) }  else { "" }
    $oldIn = if ($body) { Clean-Path ([string]$body.oldReport) } else { "" }
    if (-not $rawIn) { $rawIn = Get-DefaultRawInput  }
    if (-not $oldIn) { $oldIn = Get-DefaultOldReport }

    $script:State.params.rawInput  = $rawIn
    $script:State.params.oldReport = $oldIn
    $script:State.status           = "Running"
    $script:State.startedUtc       = (Get-Date).ToUniversalTime().ToString("o")
    $script:State.progress.msg     = "Launching..."

    Add-Log "Run started."
    Add-Log "Raw input : $rawIn"
    Add-Log "Old report: $oldIn"

    # Validate before launching -- a bad path inside the job surfaces as a
    # generic failure, whereas here we can say exactly which field is wrong.
    if (-not $rawIn -or -not (Test-Path $rawIn -ErrorAction SilentlyContinue)) {
        $script:State.status = "Failed"
        Add-Log "ERROR: raw input file not found."
        Send-Json $ctx @{ ok=$false; error="Raw input file not found:`n$rawIn" }
        return
    }
    if (-not $oldIn -or -not (Test-Path $oldIn -ErrorAction SilentlyContinue)) {
        $script:State.status = "Failed"
        Add-Log "ERROR: old Outstanding Rollover Report not found."
        Send-Json $ctx @{ ok=$false; error="Old Outstanding Rollover Report not found:`n$oldIn" }
        return
    }

    $runScript = Join-Path $AppRoot "run.ps1"
    if (-not (Test-Path $runScript)) {
        $script:State.status = "Failed"
        Add-Log "ERROR: run.ps1 not found at $runScript"
        Send-Json $ctx @{ ok=$false; error="run.ps1 not found next to server.ps1." }
        return
    }

    # Invoke run.ps1 directly in the job's runspace rather than as a nested
    # `powershell -File`. A nested process swallows early failures and exit
    # codes, so a missing module looks like a silent instant completion. Here
    # run.ps1's Write-Output lands straight in the job output stream and any
    # terminating error is re-emitted as an @@FAIL the parser understands.
    $script:Job = Start-Job -ScriptBlock {
        param($runScript, $rawIn, $oldIn)
        $ErrorActionPreference = "Stop"
        try {
            & $runScript -RawInputFile $rawIn -OldOutstandingROReport $oldIn -JobMode
        } catch {
            Write-Output "@@FAIL msg=$($_.Exception.Message)"
            Write-Output "Stack: $($_.ScriptStackTrace -replace '[\r\n]+',' | ')"
        }
    } -ArgumentList $runScript, $rawIn, $oldIn

    Send-Json $ctx @{ ok=$true }
}

function Handle-Cancel {
    param($ctx)
    if ($script:State.status -ne "Running") {
        Send-Json $ctx @{ ok=$false; error="Nothing is running." }
        return
    }
    if ($script:Job) {
        Stop-Job   -Job $script:Job -ErrorAction SilentlyContinue
        Remove-Job -Job $script:Job -Force -ErrorAction SilentlyContinue
        $script:Job = $null
    }
    $script:State.status = "Cancelled"
    Add-Log "Run cancelled by operator."
    Add-Log "NOTE: a partially written workbook may be left in the output folder."
    Send-Json $ctx @{ ok=$true }
}

# Opens the output folder (or selects the workbook) in Explorer, on the
# server's machine -- which is the same machine as the browser here.
function Handle-Open {
    param($ctx, [string]$which)
    $target = switch ($which) {
        "folder"   { $script:State.outputFiles.folder   }
        "workbook" { $script:State.outputFiles.workbook }
        default    { "" }
    }
    if (-not $target -or -not (Test-Path $target -ErrorAction SilentlyContinue)) {
        Send-Json $ctx @{ ok=$false; error="Nothing to open yet." }
        return
    }
    try {
        if ($which -eq "workbook") {
            Start-Process explorer.exe "/select,`"$target`""
        } else {
            Start-Process explorer.exe "`"$target`""
        }
        Send-Json $ctx @{ ok=$true }
    } catch {
        Send-Json $ctx @{ ok=$false; error=$_.Exception.Message }
    }
}

# ==============================================================================
# LISTENER
# ==============================================================================
$listener = [System.Net.HttpListener]::new()
$listener.Prefixes.Add("http://localhost:$Port/")
try {
    $listener.Start()
} catch {
    Write-Host "FATAL: could not start listener." -ForegroundColor Red
    Write-Host "  $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "  Check that port $Port is free (another copy may already be running)." -ForegroundColor Yellow
    Read-Host "Press Enter to exit"
    exit 1
}

$indexPath  = Join-Path $AppRoot "index.html"
$configPath = Join-Path $AppRoot "config.json"

Write-Host ""
Write-Host "=================================================" -ForegroundColor Cyan
Write-Host "  Outstanding Rollover Report Server" -ForegroundColor Cyan
Write-Host "  GUI: http://localhost:$Port/" -ForegroundColor Green
if (-not (Test-Path $configPath)) {
    Write-Host "  WARNING: config.json not found next to server.ps1" -ForegroundColor Yellow
}
Write-Host "  Close this window or use 'shut down server' in the GUI." -ForegroundColor DarkGray
Write-Host "=================================================" -ForegroundColor Cyan
Write-Host ""

$running = $true

while ($running -and $listener.IsListening) {
    $ctx    = $listener.GetContext()
    $req    = $ctx.Request
    $path   = $req.Url.AbsolutePath.TrimEnd('/')
    if ($path -eq "") { $path = "/" }
    $method = $req.HttpMethod

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
                Send-Json $ctx ([ordered]@{
                    status      = $script:State.status
                    progress    = $script:State.progress
                    phaseMax    = $PHASE_MAX
                    log         = @($script:State.log)
                    counts      = @($script:State.counts)
                    persh       = @($script:State.persh)
                    nearMiss    = @($script:State.nearMiss)
                    startedUtc  = $script:State.startedUtc
                    outputFiles = $script:State.outputFiles
                    params      = $script:State.params
                })
            }

            "^GET /defaults$" {
                $raw = Get-DefaultRawInput
                $old = Get-DefaultOldReport
                Send-Json $ctx ([ordered]@{
                    rawInput      = $raw
                    oldReport     = $old
                    rawFileDate   = Get-FileDateString $raw
                    oldFileDate   = Get-FileDateString $old
                    today         = (Get-Date).ToString("yyyy-MM-dd")
                    recent        = @(Get-RecentSpreadsheets)
                    downloads     = Get-DownloadsFolder
                    configOk      = (Test-Path $configPath)
                })
            }

            "^POST /start$"  { Handle-Start  $ctx }
            "^POST /cancel$" { Handle-Cancel $ctx }

            "^POST /openFolder$"   { Handle-Open $ctx "folder"   }
            "^POST /openWorkbook$" { Handle-Open $ctx "workbook" }

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
    Stop-Job   -Job $script:Job -ErrorAction SilentlyContinue
    Remove-Job -Job $script:Job -Force -ErrorAction SilentlyContinue
}
Write-Host "Server stopped." -ForegroundColor DarkGray
