# ============================================================================
# Outstanding Rollover Report - Generation Script (ImportExcel Rewrite)
# ============================================================================
#
# MATCHER GUARD + NEAR-MISS REVIEW (this revision):
#   - Get-TokenOverlapScore single-token anchor guard no longer compares the
#     matched token's length against $Longer[0] (the other name's FIRST
#     token). That ratio was arbitrary: the matched token is rarely the lead
#     word, so valid whole-token matches were being discarded. 'Voya' vs
#     'Reliastar Life Insurance / Voya' scored 4/9 = 0.44 and was thrown out
#     despite an exact token hit, which is why Reliastar had to be hand-added
#     as its own config entry. Replaced with a minimum token length check.
#     NOTE: single-token config entries are now greedy - any name containing
#     that word as a whole token will match. Do NOT add generic single words
#     (Trust, Retirement, Group, etc.) to redCompanies.
#   - STEP 12 now records "near misses": Blue (Contra Firm) rows whose best
#     score against a redCompanies entry landed between
#     $Script:NearMissReviewFloor and the match threshold. These print in a
#     REVIEW banner at the end of the run, next to the Pershing reminder, so
#     a firm that should have gone red does not quietly stay blue in a log
#     nobody reads. Flags are gated on sharing at least one token outside
#     $Script:MatchStopwords and $Script:WeakMatchTokens - without that gate
#     the banner filled with pairs matched only on 'retirement' / 'funds' /
#     'american' and became noise. This catches a DNC firm WORDED
#     differently; it does NOT catch a firm missing from redCompanies
#     entirely (those score near zero). Only reconciling config.json against
#     the Do Not Call resource catches those.
#
# MATCHER TUNING + DNC REMOVAL (prior revision):
#   - Test-FuzzyMatchAgainstList now evaluates ALL config entries and
#     prefers Exact > Contains > TokenOverlap, instead of returning the
#     first entry to match at any tier. Config list order no longer
#     decides attribution (fixes row 'Fidelity' token-matching
#     'American Fidelity Assurance' before reaching the exact 'Fidelity'
#     entry).
#   - Normalize-ForMatching converts separators (/ \ * & + -) to spaces
#     BEFORE stripping punctuation, so 'Services/Truist' tokenizes as
#     'services truist' instead of fusing into 'servicestruist' (fixes
#     'National Financial Services/Truist' missing the 'Truist' entry).
#   - Token-overlap scoring filters generic filler words (life, insurance,
#     co, company, financial, services, group, etc.) when both names keep
#     at least one distinctive token. Prevents score inflation between
#     unrelated companies that share only boilerplate (e.g. 'Security
#     Mutual Life Insurance Co. of NY' scoring 0.75 vs 'Principal Life
#     Insurance Co.'). Falls back to unfiltered tokens when a name is
#     all filler (so 'Principal Life Insurance' style entries still work).
#   - Single-token anchor guard made deterministic: compares against the
#     longer name's first token instead of arbitrary HashSet enumeration
#     order. (Superseded this revision - see above.)
#   - DNC Rep logic removed entirely (handled outside this script now):
#     config dnc_reps no longer read, no Notes/Callback Date stamping,
#     no gray-out formatting on the call lists.
#
# TRADING PLATFORM / PERSHING REWRITE (prior revision):
#   - NEW Step 5: "Trading Platform" classification runs FIRST (before
#     Received), highlight-only on Raw DNU. Matches Target Product Type
#     against config sheets.tradingPlatform.filters.targetProductTypes
#     (ISPLUS, OAID, PERSH, PERSHA, PERSHB, PRFP, PRFTD, SCIAR). These are
#     monitored by the Trading Platform Desk and never reach a call list.
#     Replaces both the old Pershing "Pass 1" product-type filter and the
#     retired Misc. product codes.
#   - NEW Step 10B: Pershing rows (Source Company contains the keyword from
#     config sheets.pershing.filters.sourceCompanyKeyword) are claimed AFTER
#     Pending LOA and BEFORE the call lists. Highlight-only on Raw DNU;
#     they never land on a call list automatically.
#   - The NetX check is MANUAL. At the end of the run, the script prints a
#     reminder listing every excluded Pershing row. The user checks NetX
#     and manually moves each row:
#       * LOA uploaded into NetX     -> Contra Firm Call List (blue)
#       * LOA NOT uploaded into NetX -> Do Not Call -> Advisor Call List (red)
#   - Pershing detection switched from exact match to case-insensitive
#     keyword contains so variants like "Pershing LLC" are caught.
#   - Step renumbering: Received is now Step 6, State TRSs is Step 7.
#
# SOP ALIGNMENT REWRITE (prior revision):
#   - Removed Misc. sheet entirely (Step 5)
#   - Removed Call Sheet intermediate tab (old Step 12)
#   - State TRSs, Pershing, and Internals are now Raw DNU highlight only
#     (rows still classified and marked, but no separate tab is exported)
#   - Pending LOA also converted to highlight-only on Raw DNU
#   - Advisor / Contra Firm Call Lists are built directly from unclaimed
#     rows; no intermediate Call Sheet
#   - Removed Pershing manual-review ASCII flowchart printout at end
#   - Retained LOA Date + 14 days callback auto-fill (Step 18) - quality of
#     life feature, not on SOP but kept per direction
#   - Retained NIGO carry-over from old report (Step 16B) - same reason
#
# PRIOR CHANGE LOG (retained for reference):
#   CHANGE B - Fuzzy match threshold raised from 0.6 to 0.85
#   CHANGE C - Callback dates written as DateTime objects, not strings
#   CHANGE D - NIGO rows from old report carried over if Rollover ID absent
#   CHANGE F - $Script:DebugLoggingEnabled flag gates verbose output
# ============================================================================

param(
    [Parameter(Mandatory=$true)]
    [string]$RawInputFile,

    [Parameter(Mandatory=$true)]
    [string]$OldOutstandingROReport,

    # Set by server.ps1 via Start-Job. Swaps Write-Host for Write-Output so the
    # server can drain progress live, suppresses every Read-Host, and emits the
    # @@STATUS / @@COUNT / @@PERSH / @@NEARMISS / @@DONE / @@FAIL markers the
    # server parses. Never set this when running the script by hand.
    [switch]$JobMode
)

$isJobMode = $JobMode.IsPresent

# ---------------------------------------------------------------------------
# MARKER PROTOCOL (JobMode only) - server.ps1 Parse-JobLine consumes these.
#   @@STATUS phase=N pct=N msg=text     progress bar + phase line
#   @@COUNT  <label>|<rows>             one row of the results table
#   @@PERSH  <text>                     one row needing the manual NetX check
#   @@NEARMISS <text>                   one blue company worth reviewing
#   @@DONE   workbook=.. folder=.. log=..
#   @@FAIL   msg=text
# Anything else is treated as a plain log line.
# ---------------------------------------------------------------------------
function Send-Marker {
    param([string]$Text)
    if ($isJobMode) { Write-Output $Text }
}

function Send-Status {
    param([int]$Phase, [int]$Pct, [string]$Message)
    Send-Marker "@@STATUS phase=$Phase pct=$Pct msg=$Message"
}

function Send-Count {
    param([string]$Label, [int]$Rows)
    Send-Marker "@@COUNT $Label|$Rows"
}

function Send-Fail {
    param([string]$Message)
    Send-Marker "@@FAIL msg=$Message"
}

# Replaces every bare Read-Host. In JobMode there is no console to prompt at,
# and a blocked Read-Host would hang the job until the server is killed.
function Wait-ForOperator {
    param([string]$Prompt = "Press Enter to exit")
    if (-not $isJobMode) { Read-Host $Prompt }
}

# ============================================================================
# INITIAL SETUP
# ============================================================================

$RawInputFile = $RawInputFile.Trim('"', ' ')
$OldOutstandingROReport = $OldOutstandingROReport.Trim('"', ' ')

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Send-Fail "ImportExcel module is not installed. Run: Install-Module -Name ImportExcel -Scope CurrentUser"
    if (-not $isJobMode) {
        Write-Error "ImportExcel module is not installed. Please run: Install-Module -Name ImportExcel -Scope CurrentUser"
    }
    return
}
Import-Module ImportExcel -ErrorAction Stop


# ============================================================================
# CONFIGURATION LOADING
# ============================================================================

$ConfigPath = Join-Path $ScriptDir "config.json"
if (-not (Test-Path $ConfigPath)) {
    Send-Fail "config.json not found next to run.ps1 (expected at $ConfigPath)."
    if (-not $isJobMode) {
        Write-Error "Configuration file not found at $ConfigPath. Please run the config tool first."
    }
    return
}
$Config = Get-Content $ConfigPath -Raw | ConvertFrom-Json

# --- Received Sheet Config ---
$ReceivedColor             = $Config.sheets.received.color

# --- State TRSs Config (highlight-only on Raw DNU) ---
$StateTRSCompanies         = $Config.sheets.stateTRS.filters.sourceCompanies
$StateTRSsColor            = $Config.sheets.stateTRS.color

# --- Trading Platform Config (highlight-only on Raw DNU, never called) ---
if (-not $Config.sheets.tradingPlatform) {
    Send-Fail "config.json is missing the 'sheets.tradingPlatform' section (targetProductTypes: ISPLUS, OAID, PERSH, PERSHA, PERSHB, PRFP, PRFTD, SCIAR)."
    if (-not $isJobMode) {
        Write-Error "config.json is missing the 'sheets.tradingPlatform' section. Please update config.json (targetProductTypes: ISPLUS, OAID, PERSH, PERSHA, PERSHB, PRFP, PRFTD, SCIAR)."
    }
    return
}
$TradingPlatformProductTypes = $Config.sheets.tradingPlatform.filters.targetProductTypes
$TradingPlatformColor        = $Config.sheets.tradingPlatform.color

# --- Pershing Config (highlight-only on Raw DNU; NetX check is MANUAL) ---
$PershingKeyword           = $Config.sheets.pershing.filters.sourceCompanyKeyword
if ([string]::IsNullOrWhiteSpace($PershingKeyword)) { $PershingKeyword = "Pershing" }
$PershingColor             = $Config.sheets.pershing.color
if ([string]::IsNullOrWhiteSpace($PershingColor)) { $PershingColor = "#FF7DE6" }

# --- Internals Config (highlight-only on Raw DNU) ---
$InternalsCompanies        = $Config.sheets.internals.filters.sourceCompanies
$InternalsColor            = $Config.sheets.internals.color

# --- Funding Review Sheet Config ---
$FundingReviewExcludePatterns = $Config.sheets.fundingReview.filters.targetProductExcludePatterns
$FundingReviewColor        = $Config.sheets.fundingReview.color

# --- Pending LOA Sheet Config ---
$PendingLOAColor           = $Config.sheets.pendingLOA.color

# --- Call List Config ---
$RedCompanies              = $Config.sheets.callSheet.companyColoring.redCompanies
$CallSheetRedColor         = $Config.sheets.callSheet.companyColoring.redColor
$CallSheetBlueColor        = $Config.sheets.callSheet.companyColoring.blueColor

# --- Debug logging flag ---
$Script:DebugLoggingEnabled = $true


# ============================================================================
# OUTPUT FOLDER & LOG FILE SETUP
# ============================================================================

$InputFileDirectory = Split-Path -Parent $RawInputFile
$Timestamp          = Get-Date -Format 'MM.dd.yyyy'
$TimeComponent      = Get-Date -Format 'HH-mm'
$OutputFolderName   = "$Timestamp Outstanding RO Report Generation Request $TimeComponent"
$OutputFolder       = Join-Path $InputFileDirectory $OutputFolderName

if (-not (Test-Path $OutputFolder)) {
    New-Item -Path $OutputFolder -ItemType Directory | Out-Null
}

$LogFile = Join-Path $OutputFolder "ProcessingLog_$(Get-Date -Format 'MM-dd-yyyy_HH-mm').log"


# ============================================================================
# LOGGING FUNCTION
# ============================================================================

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet('INFO','WARNING','ERROR','SUCCESS','DEBUG')]
        [string]$Level = 'INFO'
    )

    $LogTimestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $Line = "$LogTimestamp - $Level - $Message"

    if ($isJobMode) {
        # The server stamps every line it receives, so emit the message bare.
        #
        # DEBUG is deliberately withheld from the job stream. The red/blue match
        # dump alone runs to several hundred lines and would push everything
        # useful out of the server's 500-line ring buffer. It still goes to the
        # log file below, so nothing is lost.
        #
        # This suppression is also load-bearing for correctness, not just noise:
        # Test-FuzzyMatchAgainstList RETURNS a value and logs at DEBUG. If DEBUG
        # reached Write-Output here, those lines would be appended to that
        # function's output stream and every match result would come back as an
        # array instead of a hashtable. Do not make DEBUG emit in JobMode.
        if ($Level -ne 'DEBUG') {
            $Prefix = switch ($Level) {
                'ERROR'   { 'ERROR: '   }
                'WARNING' { 'WARNING: ' }
                default   { ''          }
            }
            Write-Output "$Prefix$Message"
        }
    }
    else {
        switch ($Level) {
            'INFO'    { Write-Host $Line -ForegroundColor White }
            'WARNING' { Write-Host $Line -ForegroundColor Yellow }
            'ERROR'   { Write-Host $Line -ForegroundColor Red }
            'SUCCESS' { Write-Host $Line -ForegroundColor Green }
            'DEBUG'   { Write-Host $Line -ForegroundColor DarkGray }
        }
    }

    Add-Content -Path $LogFile -Value $Line
}


# ============================================================================
# HELPER: Convert hex color string (#RRGGBB) to System.Drawing.Color
# ============================================================================

function Convert-HexToColor {
    param([string]$HexColor)

    $Hex = $HexColor.TrimStart('#')
    $R = [Convert]::ToInt32($Hex.Substring(0, 2), 16)
    $G = [Convert]::ToInt32($Hex.Substring(2, 2), 16)
    $B = [Convert]::ToInt32($Hex.Substring(4, 2), 16)

    return [System.Drawing.Color]::FromArgb($R, $G, $B)
}


# ============================================================================
# FUZZY STRING MATCHING
# ============================================================================

$Script:FuzzyMatchThreshold = 0.85

# Minimum length for a single-token name to be allowed to claim a longer
# name via token overlap. See the anchor guard in Get-TokenOverlapScore.
# 3 keeps short real firm names (UBS) working; anything shorter is noise.
$Script:MinAnchorTokenLength = 3

# Blue (Contra Firm) rows scoring at or above this against a redCompanies
# entry - but below $Script:FuzzyMatchThreshold - are flagged for manual
# review at the end of the run. Lower this to widen the review net.
$Script:NearMissReviewFloor = 0.5

# Generic filler words that inflate token-overlap scores between unrelated
# companies. Filtered out before scoring, but ONLY when both names keep at
# least one distinctive token after filtering (see Get-TokenOverlapScore).
# Deliberately NOT included: funds, fund, annuities, retirement, bank,
# mutual, national - removing those would silently flip real match outcomes
# (e.g. 'John Hancock Annuities' would start matching 'John Hancock Funds').
$Script:MatchStopwords = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]@('co','company','companies','inc','incorporated','llc','corp',
                'corporation','of','the','and','life','insurance','ins',
                'financial','services','service','group'),
    [System.StringComparer]::OrdinalIgnoreCase
)

# Tokens distinctive enough to affect scoring but far too common across
# firm names to justify a manual review flag on their own. Used ONLY by the
# near-miss review report - these do NOT affect matching. Without this
# filter the review banner fills with pairs matched purely on 'retirement',
# 'funds', 'american', 'wealth' etc. and stops getting read. Measured on a
# live call list: 17 flags without it, 0 with it, and a genuinely reworded
# DNC firm ('Schwab Retirement Plan Services') still flags on 'schwab'.
$Script:WeakMatchTokens = [System.Collections.Generic.HashSet[string]]::new(
    [string[]]@('american','america','national','nationwide','new','york','us','usa',
                'retirement','retirements','fund','funds','benefit','benefits',
                'employee','employees','wealth','bank','banking','mutual','trust',
                'annuity','annuities','plan','plans','savings','investment',
                'investments','capital','management','teachers','securities'),
    [System.StringComparer]::OrdinalIgnoreCase
)

function Normalize-ForMatching {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return "" }
    $Cleaned = $Value.ToLower().Trim()
    # Separators become spaces FIRST so glued names split into real tokens
    # ('Services/Truist' -> 'services truist', not 'servicestruist').
    # Hyphens too: 'TIAA-CREF' and 'TIAA Cref' now normalize identically.
    $Cleaned = $Cleaned -replace '[/\\*&+\-]', ' '
    $Cleaned = $Cleaned -replace '[^\w\s]', ''
    $Cleaned = $Cleaned -replace '\s+', ' '
    return $Cleaned.Trim()
}

function Get-TokenList {
    param([string]$NormalizedValue)
    # Ordered, de-duplicated token list.
    $Tokens = [System.Collections.Generic.List[string]]::new()
    if ([string]::IsNullOrWhiteSpace($NormalizedValue)) { return ,$Tokens }
    foreach ($Token in ($NormalizedValue -split '\s+')) {
        if (-not [string]::IsNullOrWhiteSpace($Token) -and -not $Tokens.Contains($Token)) {
            [void]$Tokens.Add($Token)
        }
    }
    return ,$Tokens
}

function Remove-StopwordTokens {
    param([System.Collections.Generic.List[string]]$Tokens)
    $Filtered = [System.Collections.Generic.List[string]]::new()
    foreach ($Token in $Tokens) {
        if (-not $Script:MatchStopwords.Contains($Token)) { [void]$Filtered.Add($Token) }
    }
    return ,$Filtered
}

function Get-SharedDistinctiveTokens {
    # Tokens two names genuinely share, ignoring filler and weak tokens.
    # Used to decide whether a near miss is worth a human's attention.
    param(
        [string]$ValueA,
        [string]$ValueB
    )
    $TokensA = Get-TokenList -NormalizedValue (Normalize-ForMatching -Value $ValueA)
    $TokensB = Get-TokenList -NormalizedValue (Normalize-ForMatching -Value $ValueB)

    $Shared = [System.Collections.Generic.List[string]]::new()
    foreach ($Token in $TokensA) {
        if (-not $TokensB.Contains($Token)) { continue }
        if ($Script:MatchStopwords.Contains($Token)) { continue }
        if ($Script:WeakMatchTokens.Contains($Token)) { continue }
        [void]$Shared.Add($Token)
    }
    return ,$Shared
}

function Get-TokenOverlapScore {
    param(
        [System.Collections.Generic.List[string]]$TokensA,
        [System.Collections.Generic.List[string]]$TokensB
    )
    if ($TokensA.Count -eq 0 -or $TokensB.Count -eq 0) { return 0.0 }

    # Drop generic filler tokens, but only if BOTH sides keep at least one
    # distinctive token. Otherwise fall back to the unfiltered lists so
    # all-filler names (e.g. 'Principal Life Insurance') can still match.
    $FilteredA = Remove-StopwordTokens -Tokens $TokensA
    $FilteredB = Remove-StopwordTokens -Tokens $TokensB
    if ($FilteredA.Count -gt 0 -and $FilteredB.Count -gt 0) {
        $TokensA = $FilteredA
        $TokensB = $FilteredB
    }

    # NOTE: direct assignments on purpose. Assigning a collection out of an
    # if-EXPRESSION unrolls it, so a single-token list collapses to a bare
    # string and $Shorter[0] silently becomes a char.
    if ($TokensA.Count -le $TokensB.Count) {
        $Shorter = $TokensA
        $Longer  = $TokensB
    }
    else {
        $Shorter = $TokensB
        $Longer  = $TokensA
    }

    $SharedCount = 0
    foreach ($Token in $Shorter) {
        if ($Longer.Contains($Token)) { $SharedCount++ }
    }
    $RawScore = [double]$SharedCount / [double]$Shorter.Count

    # Single-token anchor guard: a one-token name may claim a longer name
    # only if that token is distinctive enough on its own.
    #
    # The old test compared $Shorter[0].Length against $Longer[0].Length -
    # the other name's FIRST token - which is arbitrary, because the matched
    # token is usually not the lead word. 'Voya' vs 'Reliastar Life
    # Insurance / Voya' scored 4/9 = 0.44 and was discarded despite an exact
    # whole-token hit. A token-overlap match is already whole-token exact,
    # so substring-fragment risk does not apply here; length alone is the
    # only thing worth testing.
    #
    # TRADE-OFF: single-token config entries are now greedy. Any Source
    # Company containing that word as a whole token will match. That is the
    # intent for firm names, but do not add generic single words to
    # redCompanies.
    if ($RawScore -ge 1.0 -and $Shorter.Count -eq 1) {
        if ($Shorter[0].Length -lt $Script:MinAnchorTokenLength) { return 0.0 }
    }

    return $RawScore
}

function Test-FuzzyMatch {
    param(
        [string]$DataValue,
        [string]$ConfigValue,
        [string]$Label = "Match",
        [double]$Threshold = $Script:FuzzyMatchThreshold
    )
    $NormData   = Normalize-ForMatching -Value $DataValue
    $NormConfig = Normalize-ForMatching -Value $ConfigValue
    if ($NormData -eq $NormConfig) {
        return @{ IsMatch = $true; Tier = "Exact"; Score = 1.0 }
    }
    if ($NormData.Contains($NormConfig) -or $NormConfig.Contains($NormData)) {
        $LenA = $NormData.Length
        $LenB = $NormConfig.Length
        $LengthRatio = [double][Math]::Min($LenA, $LenB) / [double][Math]::Max($LenA, $LenB)
        if ($LengthRatio -ge 0.65) {
            return @{ IsMatch = $true; Tier = "Contains"; Score = 1.0 }
        }
    }
    $TokensData   = Get-TokenList -NormalizedValue $NormData
    $TokensConfig = Get-TokenList -NormalizedValue $NormConfig
    $Score        = Get-TokenOverlapScore -TokensA $TokensData -TokensB $TokensConfig
    if ($Score -ge $Threshold) {
        return @{ IsMatch = $true; Tier = "TokenOverlap"; Score = $Score }
    }
    return @{ IsMatch = $false; Tier = "None"; Score = $Score }
}

function Test-FuzzyMatchAgainstList {
    param(
        [string]$DataValue,
        [array]$ConfigValues,
        [string]$Label = "Match",
        [double]$Threshold = $Script:FuzzyMatchThreshold
    )
    if ([string]::IsNullOrWhiteSpace($DataValue) -or $ConfigValues.Count -eq 0) {
        return @{ IsMatch = $false; MatchedConfig = $null; Tier = "None"; Score = 0.0 }
    }

    # Evaluate EVERY config entry, then prefer matches by tier quality:
    # Exact > Contains > TokenOverlap. Previously the first entry to match
    # at ANY tier won, so config list order could beat match quality.
    # Within the same tier, highest score wins; ties keep the first entry.
    $TierRank = @{ "Exact" = 3; "Contains" = 2; "TokenOverlap" = 1 }

    $BestMatch   = $null
    $BestNoMatch = @{ IsMatch = $false; MatchedConfig = $null; Tier = "None"; Score = 0.0 }

    foreach ($ConfigValue in $ConfigValues) {
        $Result = Test-FuzzyMatch -DataValue $DataValue -ConfigValue $ConfigValue -Label $Label -Threshold $Threshold

        if ($Result.IsMatch) {
            $Rank = $TierRank[$Result.Tier]
            if ($null -eq $BestMatch -or
                $Rank -gt $BestMatch.Rank -or
                ($Rank -eq $BestMatch.Rank -and $Result.Score -gt $BestMatch.Score)) {
                $BestMatch = @{ MatchedConfig = $ConfigValue; Tier = $Result.Tier; Score = $Result.Score; Rank = $Rank }
            }
            if ($Rank -eq 3) { break }   # Exact cannot be beaten - stop early
        }
        elseif ($Result.Score -gt $BestNoMatch.Score) {
            $BestNoMatch = @{ IsMatch = $false; MatchedConfig = $ConfigValue; Tier = "None"; Score = $Result.Score }
        }
    }

    if ($null -ne $BestMatch) {
        if ($Script:DebugLoggingEnabled) {
            Write-Log "    [$Label] MATCH    '$DataValue'  ->  '$($BestMatch.MatchedConfig)'  (method: $($BestMatch.Tier), score: $([math]::Round($BestMatch.Score, 2)))" -Level DEBUG
        }
        return @{ IsMatch = $true; MatchedConfig = $BestMatch.MatchedConfig; Tier = $BestMatch.Tier; Score = $BestMatch.Score }
    }

    if ($Script:DebugLoggingEnabled -and $BestNoMatch.Score -gt 0.3) {
        Write-Log "    [$Label] NO MATCH '$DataValue'  (closest: '$($BestNoMatch.MatchedConfig)', score: $([math]::Round($BestNoMatch.Score, 2)) - threshold: $Threshold)" -Level DEBUG
    }
    return $BestNoMatch
}


# ============================================================================
# OLD REPORT READER
# ============================================================================

function Read-OldReportSheetData {
    # Fills the caller-supplied -Into hashtable instead of returning one.
    #
    # This function logs, and in JobMode Write-Log emits via Write-Output. A
    # function that both logs and returns a value would have its log lines
    # appended to its own output stream, so the caller would receive an array
    # of [log strings + hashtable] rather than the hashtable. That would break
    # the notes/callback carry-over silently - every old-report value would be
    # dropped and nobody would notice until someone missed a callback. Taking
    # the destination by reference sidesteps the whole class of problem.
    param(
        [string]$FilePath,
        [string]$SheetName,
        [Parameter(Mandatory=$true)][hashtable]$Into,
        [switch]$IncludeCallbackDate,
        [switch]$IncludeSourceCompanyFormatting
    )

    Write-Log "  Reading old report data from sheet '$SheetName'..."

    try {
        $SheetNames = (Get-ExcelSheetInfo -Path $FilePath).Name
    }
    catch {
        Write-Log "  Could not read sheet info from old report: $($_.Exception.Message)" -Level WARNING
        return
    }

    if ($SheetName -notin $SheetNames) {
        Write-Log "  Sheet '$SheetName' not found in old report - skipping" -Level WARNING
        return
    }

    $ExcelPackage = Open-ExcelPackage -Path $FilePath

    try {
        $Worksheet = $ExcelPackage.Workbook.Worksheets[$SheetName]

        if ($null -eq $Worksheet) {
            Write-Log "  Worksheet '$SheetName' is null after opening - skipping" -Level WARNING
            return
        }

        $StartRow = $Worksheet.Dimension.Start.Row
        $EndRow   = $Worksheet.Dimension.End.Row
        $EndCol   = $Worksheet.Dimension.End.Column

        $HeaderMap = @{}
        for ($col = 1; $col -le $EndCol; $col++) {
            $HeaderText = $Worksheet.Cells[$StartRow, $col].Text.Trim()
            if (-not [string]::IsNullOrWhiteSpace($HeaderText)) {
                $HeaderMap[$HeaderText.ToLower()] = $col
            }
        }

        $RolloverIDCol = $HeaderMap["rollover id"]
        $NotesCol      = $HeaderMap["notes"]
        $CallbackCol   = if ($IncludeCallbackDate) { $HeaderMap["callback date"] } else { $null }

        if (-not $RolloverIDCol) {
            Write-Log "  'Rollover ID' column not found in '$SheetName'" -Level WARNING
            return
        }

        $DataCount = 0
        for ($row = ($StartRow + 1); $row -le $EndRow; $row++) {
            $RolloverID = $Worksheet.Cells[$row, $RolloverIDCol].Text.Trim()
            if ([string]::IsNullOrWhiteSpace($RolloverID)) { continue }

            $Entry = @{}

            if ($NotesCol) {
                $NotesCell = $Worksheet.Cells[$row, $NotesCol]
                $Entry.Notes          = $NotesCell.Text
                $Entry.NotesBold      = $NotesCell.Style.Font.Bold
                $Entry.NotesItalic    = $NotesCell.Style.Font.Italic
                $Entry.NotesFontColor = $null
                $Entry.NotesBgColor   = $null

                try {
                    $FontColorRgb = $NotesCell.Style.Font.Color.Rgb
                    if (-not [string]::IsNullOrWhiteSpace($FontColorRgb)) { $Entry.NotesFontColor = $FontColorRgb }
                } catch { }

                try {
                    if ($NotesCell.Style.Fill.PatternType -ne [OfficeOpenXml.Style.ExcelFillStyle]::None) {
                        $BgRgb = $NotesCell.Style.Fill.BackgroundColor.Rgb
                        if (-not [string]::IsNullOrWhiteSpace($BgRgb)) { $Entry.NotesBgColor = $BgRgb }
                    }
                } catch { }
            }

            if ($CallbackCol) {
                $CallbackCell = $Worksheet.Cells[$row, $CallbackCol]
                $Entry.CallbackDate          = $CallbackCell.Text
                $Entry.CallbackDateBold      = $CallbackCell.Style.Font.Bold
                $Entry.CallbackDateItalic    = $CallbackCell.Style.Font.Italic
                $Entry.CallbackDateFontColor = $null
                $Entry.CallbackDateBgColor   = $null
                $Entry.CallbackDateRawValue  = $CallbackCell.Value

                try {
                    $FontColorRgb = $CallbackCell.Style.Font.Color.Rgb
                    if (-not [string]::IsNullOrWhiteSpace($FontColorRgb)) { $Entry.CallbackDateFontColor = $FontColorRgb }
                } catch { }

                try {
                    if ($CallbackCell.Style.Fill.PatternType -ne [OfficeOpenXml.Style.ExcelFillStyle]::None) {
                        $BgRgb = $CallbackCell.Style.Fill.BackgroundColor.Rgb
                        if (-not [string]::IsNullOrWhiteSpace($BgRgb)) { $Entry.CallbackDateBgColor = $BgRgb }
                    }
                } catch { }
            }

            if ($IncludeSourceCompanyFormatting) {
                $SrcCompCol = $HeaderMap["source company"]
                if ($SrcCompCol) {
                    $SrcCell = $Worksheet.Cells[$row, $SrcCompCol]
                    $Entry.SourceCompanyBgColor = $null
                    try {
                        if ($SrcCell.Style.Fill.PatternType -ne [OfficeOpenXml.Style.ExcelFillStyle]::None) {
                            $BgRgb = $SrcCell.Style.Fill.BackgroundColor.Rgb
                            if (-not [string]::IsNullOrWhiteSpace($BgRgb)) { $Entry.SourceCompanyBgColor = $BgRgb }
                        }
                    } catch { }
                }
            }

            $HasNotes    = -not [string]::IsNullOrWhiteSpace($Entry.Notes)
            $HasCallback = $IncludeCallbackDate -and -not [string]::IsNullOrWhiteSpace($Entry.CallbackDate)

            if ($HasNotes -or $HasCallback) {
                $Into[$RolloverID] = $Entry
                $DataCount++
            }
        }

        Write-Log "  Read $DataCount entries with data from '$SheetName'" -Level SUCCESS
    }
    finally {
        $ExcelPackage.Dispose()
    }
}


# ============================================================================
# HELPER: Apply old report data to a worksheet
# ============================================================================

function Apply-OldDataToSheet {
    param(
        [OfficeOpenXml.ExcelWorksheet]$Worksheet,
        [hashtable]$OldData,
        [string]$SheetName,
        [switch]$IncludeCallbackDate,
        [switch]$IncludeSourceCompanyFormatting
    )

    if ($OldData.Count -eq 0) {
        Write-Log "  No old data to apply to '$SheetName' - skipping"
        return
    }

    Write-Log "  Applying old data to '$SheetName'..."

    $EndRow = $Worksheet.Dimension.End.Row
    $EndCol = $Worksheet.Dimension.End.Column

    $HeaderMap = @{}
    for ($col = 1; $col -le $EndCol; $col++) {
        $HeaderText = $Worksheet.Cells[1, $col].Text.Trim()
        if (-not [string]::IsNullOrWhiteSpace($HeaderText)) {
            $HeaderMap[$HeaderText.ToLower()] = $col
        }
    }

    $RolloverIDCol = $HeaderMap["rollover id"]
    $NotesCol      = $HeaderMap["notes"]
    $CallbackCol   = if ($IncludeCallbackDate) { $HeaderMap["callback date"] } else { $null }
    $SrcCompCol    = if ($IncludeSourceCompanyFormatting) { $HeaderMap["source company"] } else { $null }

    if (-not $RolloverIDCol) {
        Write-Log "  'Rollover ID' column not found in '$SheetName'" -Level WARNING
        return
    }

    $MatchCount = 0
    $NewCount   = 0

    for ($row = 2; $row -le $EndRow; $row++) {
        $RolloverID = $Worksheet.Cells[$row, $RolloverIDCol].Text.Trim()
        if ([string]::IsNullOrWhiteSpace($RolloverID)) { continue }

        if ($OldData.ContainsKey($RolloverID)) {
            $MatchCount++
            $Entry = $OldData[$RolloverID]

            if ($NotesCol -and -not [string]::IsNullOrWhiteSpace($Entry.Notes)) {
                $Cell = $Worksheet.Cells[$row, $NotesCol]
                $Cell.Value = $Entry.Notes
                if ($Entry.NotesBold)   { $Cell.Style.Font.Bold   = $true }
                if ($Entry.NotesItalic) { $Cell.Style.Font.Italic = $true }
                if ($Entry.NotesFontColor) {
                    try { $Cell.Style.Font.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml("#$($Entry.NotesFontColor -replace '^FF','')")) } catch { }
                }
                if ($Entry.NotesBgColor) {
                    try {
                        $Cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                        $Cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#$($Entry.NotesBgColor -replace '^FF','')"))
                    } catch { }
                }
            }

            if ($CallbackCol -and -not [string]::IsNullOrWhiteSpace($Entry.CallbackDate)) {
                $Cell = $Worksheet.Cells[$row, $CallbackCol]

                $WrittenAsDate = $false

                if ($Entry.CallbackDateRawValue -is [DateTime]) {
                    $Cell.Value = $Entry.CallbackDateRawValue
                    $Cell.Style.Numberformat.Format = "mm/dd/yyyy"
                    $WrittenAsDate = $true
                }
                elseif (-not $WrittenAsDate) {
                    $ParsedDate = [DateTime]::MinValue
                    if ([DateTime]::TryParse($Entry.CallbackDate, [ref]$ParsedDate)) {
                        $Cell.Value = $ParsedDate
                        $Cell.Style.Numberformat.Format = "mm/dd/yyyy"
                        $WrittenAsDate = $true
                    }
                }

                if (-not $WrittenAsDate) {
                    $Cell.Value = $Entry.CallbackDate
                }

                if ($Entry.CallbackDateBold)   { $Cell.Style.Font.Bold   = $true }
                if ($Entry.CallbackDateItalic) { $Cell.Style.Font.Italic = $true }
                if ($Entry.CallbackDateFontColor) {
                    try { $Cell.Style.Font.Color.SetColor([System.Drawing.ColorTranslator]::FromHtml("#$($Entry.CallbackDateFontColor -replace '^FF','')")) } catch { }
                }
                if ($Entry.CallbackDateBgColor) {
                    try {
                        $Cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                        $Cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#$($Entry.CallbackDateBgColor -replace '^FF','')"))
                    } catch { }
                }
            }

            if ($SrcCompCol -and $Entry.SourceCompanyBgColor) {
                try {
                    $Cell = $Worksheet.Cells[$row, $SrcCompCol]
                    $Cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $Cell.Style.Fill.BackgroundColor.SetColor([System.Drawing.ColorTranslator]::FromHtml("#$($Entry.SourceCompanyBgColor -replace '^FF','')"))
                } catch { }
            }
        }
        else {
            $NewCount++
        }
    }

    Write-Log "    Matched: $MatchCount | New (no old data): $NewCount" -Level SUCCESS
}


# ============================================================================
# HELPER: Export a set of rows to a named worksheet
# ============================================================================

function Export-SheetData {
    param(
        [string]$Path,
        [string]$WorksheetName,
        [array]$Data,
        [string[]]$ExcludeProperties = @("_ProcessedBy","_IsRedCompany"),
        [switch]$IncludeNotesColumn,
        [switch]$IncludeCallbackDateColumn
    )

    if ($Data.Count -eq 0) {
        Write-Log "  No data to export for '$WorksheetName' - creating empty sheet" -Level WARNING
        $EmptyRow = [PSCustomObject]@{ "No Data" = "No matching records found" }
        $EmptyRow | Export-Excel -Path $Path -WorksheetName $WorksheetName -ClearSheet
        return
    }

    Write-Log "  Exporting $($Data.Count) rows to '$WorksheetName'..."

    $AllProperties = $Data[0].PSObject.Properties.Name | Where-Object { $_ -notin $ExcludeProperties }

    $Data |
        Select-Object -Property $AllProperties |
        Export-Excel -Path $Path `
                     -WorksheetName $WorksheetName `
                     -ClearSheet `
                     -AutoSize `
                     -AutoFilter `
                     -FreezeTopRow `
                     -BoldTopRow

    Write-Log "  '$WorksheetName' exported successfully" -Level SUCCESS
}


# ============================================================================
# MAIN PROCESSING LOGIC
# ============================================================================

try {

    Write-Log "========================================"
    Write-Log "Rollover Processing Started"
    Write-Log "========================================"
    Write-Log "Raw Input File: $RawInputFile"
    Write-Log "Old Outstanding Report: $OldOutstandingROReport"
    Write-Log "Output Folder: $OutputFolder"
    Write-Log ""

    if (-not (Test-Path $RawInputFile)) {
        Write-Log "Input file not found: $RawInputFile" -Level ERROR
        Send-Fail "Input file not found: $RawInputFile"
        Wait-ForOperator
        exit 1
    }

    # Named from the same $Timestamp the output folder uses, so the file and
    # the folder it sits in always show the same date. Running twice in one day
    # is fine - the folder name carries the time component, so the full paths
    # still differ.
    $OutputFileName = "$Timestamp Outstanding Rollover Report.xlsx"
    $OutputFilePath = Join-Path $OutputFolder $OutputFileName

    Write-Log "Copying input file to output location..."
    Copy-Item -Path $RawInputFile -Destination $OutputFilePath -Force
    Write-Log "Output file: $OutputFilePath"
    Write-Log ""


    # ====================================================================
    # STEP 1: READ RAW DATA INTO MEMORY
    # ====================================================================

    Send-Status 1 4 "Reading raw data"

    Write-Log "========================================"
    Write-Log "STEP 1: Reading Raw Data"
    Write-Log "========================================"
    Write-Log ""

    $RawImport = Import-Excel -Path $OutputFilePath -NoHeader -StartRow 1

    $FirstCellValue = $null
    if ($RawImport.Count -gt 0) {
        $FirstCellValue = ($RawImport[0].PSObject.Properties | Select-Object -First 1).Value
    }

    if ($FirstCellValue -eq "Rollover Report") {
        Write-Log "  Detected title row 'Rollover Report' - importing from row 2..."
        $RawData = Import-Excel -Path $OutputFilePath -StartRow 2
    }
    else {
        Write-Log "  No title row detected - importing from row 1..."
        $RawData = Import-Excel -Path $OutputFilePath
    }

    Write-Log "  Imported $($RawData.Count) rows from raw report"

    if ($RawData.Count -eq 0) {
        Write-Log "No data found in raw input file!" -Level ERROR
        Send-Fail "No data rows found in the raw input file."
        Wait-ForOperator
        exit 1
    }

    $DetectedHeaders = $RawData[0].PSObject.Properties.Name
    if ($Script:DebugLoggingEnabled) { Write-Log "  Detected columns: $($DetectedHeaders -join ', ')" }
    Write-Log ""


    # ====================================================================
    # STEP 2: ADD INTERNAL TRACKING AND NOTES COLUMNS
    # ====================================================================

    Write-Log "========================================"
    Write-Log "STEP 2: Preparing Data (Adding Notes & Tracking Columns)"
    Write-Log "========================================"
    Write-Log ""

    foreach ($Row in $RawData) {
        if (-not ($Row.PSObject.Properties.Name -contains "Notes")) {
            $Row | Add-Member -NotePropertyName "Notes" -NotePropertyValue "" -Force
        }
        $Row | Add-Member -NotePropertyName "_ProcessedBy" -NotePropertyValue $null -Force
    }

    Write-Log "  Added '_ProcessedBy' tracking column to all $($RawData.Count) rows"


    # ====================================================================
    # STEP 3: SORT DATA
    # ====================================================================

    Write-Log "========================================"
    Write-Log "STEP 3: Sorting Data"
    Write-Log "========================================"
    Write-Log ""

    $RawData = $RawData | Sort-Object "Source Company", "Client Name"

    Write-Log "  Data sorted by 'Source Company', then 'Client Name'" -Level SUCCESS
    Write-Log ""


    # ====================================================================
    # STEP 4: DEFINE DESIRED COLUMN ORDER
    # ====================================================================

    $DesiredColumnOrder = @(
        "Source Company",
        "Client Name",
        "Notes",
        "Source Account",
        "OES",
        "Rollover ID",
        "Representative",
        "LOA Date",
        "TPA Authorization Requested",
        "TPA Approval",
        "Expected Amount",
        "Received Amount",
        "Target Account",
        "Target Product Type",
        "Target Product",
        "Status",
        "Client Initiated",
        "Internal",
        "Funded at a Later Date"
    )

    $ActualProperties = $RawData[0].PSObject.Properties.Name | Where-Object { $_ -ne "_ProcessedBy" }
    $OrderedProperties = @()

    foreach ($Col in $DesiredColumnOrder) {
        if ($Col -in $ActualProperties) { $OrderedProperties += $Col }
    }
    foreach ($Col in $ActualProperties) {
        if ($Col -notin $OrderedProperties) { $OrderedProperties += $Col }
    }

    Write-Log "  Column order established: $($OrderedProperties -join ', ')"
    Write-Log ""


    # ====================================================================
    # STEP 5: CLASSIFY TRADING PLATFORM ROWS (RAW DNU HIGHLIGHT ONLY)
    # Per SOP, this runs FIRST. We do not call on Trading Platform
    # products - these RO's are monitored by the Trading Platform Desk.
    # ====================================================================

    Send-Status 2 15 "Classifying Trading Platform rows"

    Write-Log "========================================"
    Write-Log "STEP 5: Classifying Trading Platform Rows (highlight-only on Raw DNU)"
    Write-Log "========================================"
    Write-Log ""

    $TradingPlatformRows = $RawData | Where-Object {
        ($null -eq $_._ProcessedBy) -and
        ($TradingPlatformProductTypes -icontains $_."Target Product Type")
    }

    Write-Log "  Found $($TradingPlatformRows.Count) rows with Target Product Type in: $($TradingPlatformProductTypes -join ', ')"

    foreach ($Row in $TradingPlatformRows) {
        $Row._ProcessedBy = "Trading Platform"
    }
    Send-Count "Trading Platform (not called)" @($TradingPlatformRows).Count

    Write-Log "  (Rows marked for Raw DNU highlight - no separate tab exported, no call list)" -Level SUCCESS
    Write-Log ""


    # ====================================================================
    # STEP 6: CLASSIFY RECEIVED ROWS (EXPORTED)
    # ====================================================================

    Send-Status 2 20 "Building Received sheet"

    Write-Log "========================================"
    Write-Log "STEP 6: Processing Received Sheet"
    Write-Log "========================================"
    Write-Log ""

    $ReceivedRows = $RawData | Where-Object {
        ($null -eq $_._ProcessedBy) -and
        (-not [string]::IsNullOrWhiteSpace($_."Received Amount")) -and
        ($_."Received Amount" -ne "0.00  $") -and
        ($_."Received Amount" -ne "0.00") -and
        ($_."Received Amount" -ne "0") -and
        ($_."Received Amount" -ne "$0.00") -and
        ($_."Received Amount" -ne 0)
    }

    Write-Log "  Found $($ReceivedRows.Count) rows with non-zero Received Amount"

    foreach ($Row in $ReceivedRows) { $Row._ProcessedBy = "Received" }
    Send-Count "Received" @($ReceivedRows).Count

    Export-SheetData -Path $OutputFilePath -WorksheetName "Received" -Data ($ReceivedRows | Select-Object -Property $OrderedProperties)

    Write-Log ""


    # ====================================================================
    # STEP 7: CLASSIFY STATE TRSs ROWS (RAW DNU HIGHLIGHT ONLY)
    # ====================================================================

    Send-Status 2 26 "Classifying State TRS rows"

    Write-Log "========================================"
    Write-Log "STEP 7: Classifying State TRSs Rows (highlight-only on Raw DNU)"
    Write-Log "========================================"
    Write-Log ""

    $StateTRSHardcoded = @("ERS", "ERS Texas", "PSERS", "SERS", "TRS", "TRS Texas", "PARS", "FRS", "DRS")
    $StateTRSBlacklist  = @("VOYA")

    $StateTRSRows = $RawData | Where-Object {
        if ($null -ne $_._ProcessedBy) { return $false }
        $SourceCompany = $_."Source Company"
        if ([string]::IsNullOrWhiteSpace($SourceCompany)) { return $false }

        if ($StateTRSBlacklist -icontains $SourceCompany.Trim()) { return $false }
        if ($StateTRSHardcoded -icontains $SourceCompany.Trim()) { return $true }

        $MatchResult = Test-FuzzyMatchAgainstList `
            -DataValue $SourceCompany `
            -ConfigValues $StateTRSCompanies `
            -Label "StateTRS" `
            -Threshold 0.98
        return $MatchResult.IsMatch
    }

    Write-Log "  Found $($StateTRSRows.Count) rows matching State TRS companies"

    foreach ($Row in $StateTRSRows) { $Row._ProcessedBy = "State TRSs" }
    Send-Count "State TRSs (not called)" @($StateTRSRows).Count

    Write-Log "  (Rows marked for Raw DNU highlight - no separate tab exported)" -Level SUCCESS
    Write-Log ""


    # ====================================================================
    # STEP 8: CLASSIFY INTERNALS ROWS (RAW DNU HIGHLIGHT ONLY)
    # ====================================================================

    Send-Status 2 30 "Classifying internal rows"

    Write-Log "========================================"
    Write-Log "STEP 8: Classifying Internals Rows (highlight-only on Raw DNU)"
    Write-Log "========================================"
    Write-Log ""

    $InternalsRows = $RawData | Where-Object {
        if ($null -ne $_._ProcessedBy) { return $false }
        $SourceCompany = $_."Source Company"
        if ([string]::IsNullOrWhiteSpace($SourceCompany)) { return $false }
        $MatchResult = Test-FuzzyMatchAgainstList `
            -DataValue $SourceCompany `
            -ConfigValues $InternalsCompanies `
            -Label "Internals"
        return $MatchResult.IsMatch
    }

    Write-Log "  Found $($InternalsRows.Count) rows matching internal companies"

    foreach ($Row in $InternalsRows) { $Row._ProcessedBy = "Internals" }
    Send-Count "Internals (not called)" @($InternalsRows).Count

    Write-Log "  (Rows marked for Raw DNU highlight - no separate tab exported)" -Level SUCCESS
    Write-Log ""


    # ====================================================================
    # STEP 9: PROCESS FUNDING REVIEW SHEET (EXPORTED)
    # ====================================================================

    Send-Status 2 35 "Building Funding Review sheet"

    Write-Log "========================================"
    Write-Log "STEP 9: Processing Funding Review Sheet"
    Write-Log "========================================"
    Write-Log ""

    $FundingReviewRows = $RawData | Where-Object {
        if ($null -ne $_._ProcessedBy) { return $false }

        $LOABlank    = [string]::IsNullOrWhiteSpace($_."LOA Date")
        $TPAReqBlank = [string]::IsNullOrWhiteSpace($_."TPA Authorization Requested")
        $TPAAppBlank = [string]::IsNullOrWhiteSpace($_."TPA Approval")
        $StatusBlank = [string]::IsNullOrWhiteSpace($_."Status")

        if (-not ($LOABlank -and $TPAReqBlank -and $TPAAppBlank -and $StatusBlank)) { return $false }

        $TargetProduct = $_."Target Product"
        if ([string]::IsNullOrWhiteSpace($TargetProduct)) { return $true }

        $IsExcluded = $false
        foreach ($Pattern in $FundingReviewExcludePatterns) {
            $RegexPattern = "^" + [regex]::Escape($Pattern).Replace("\*", ".*") + "$"
            if ($TargetProduct -match $RegexPattern) { $IsExcluded = $true; break }
        }

        return (-not $IsExcluded)
    }

    Write-Log "  Found $($FundingReviewRows.Count) rows for Funding Review"

    foreach ($Row in $FundingReviewRows) { $Row._ProcessedBy = "Funding Review" }
    Send-Count "Funding Review" @($FundingReviewRows).Count

    Export-SheetData -Path $OutputFilePath -WorksheetName "Funding Review" -Data ($FundingReviewRows | Select-Object -Property $OrderedProperties)

    Write-Log ""





    # ====================================================================
    # STEP 10: CLASSIFY PENDING LOA ROWS (RAW DNU HIGHLIGHT ONLY)
    # ====================================================================

    Send-Status 2 40 "Classifying Pending LOA rows"

    Write-Log "========================================"
    Write-Log "STEP 10: Classifying Pending LOA Rows (highlight-only on Raw DNU)"
    Write-Log "========================================"
    Write-Log ""

    $PendingLOARows = $RawData | Where-Object {
        ($null -eq $_._ProcessedBy) -and
        ([string]::IsNullOrWhiteSpace($_."LOA Date")) -and
        ($_."Status" -ine "Transfer In IGO - LOA Doc Sent")
    }

    Write-Log "  Found $($PendingLOARows.Count) rows for Pending LOA (excluding 'Transfer In IGO - LOA Doc Sent')"

    foreach ($Row in $PendingLOARows) { $Row._ProcessedBy = "Pending LOA" }
    Send-Count "Pending LOA (not called)" @($PendingLOARows).Count

    Write-Log "  (Rows marked for Raw DNU highlight - no separate tab exported)" -Level SUCCESS
    Write-Log ""


    # ====================================================================
    # STEP 10B: CLASSIFY PERSHING ROWS (RAW DNU HIGHLIGHT ONLY)
    # Pershing rows are filtered OUT before the call lists. The NetX check
    # is a MANUAL step - a reminder listing these rows prints at the end
    # of the run. The user checks NetX and manually moves each row:
    #   LOA uploaded into NetX     -> Contra Firm Call List
    #   LOA NOT uploaded into NetX -> Do Not Call -> Advisor Call List
    # Runs AFTER Pending LOA so this list matches the rows that would
    # otherwise have reached the call lists (mirrors the manual SOP order).
    # ====================================================================

    Send-Status 2 44 "Classifying Pershing rows"

    Write-Log "========================================"
    Write-Log "STEP 10B: Classifying Pershing Rows (highlight-only, manual NetX check)"
    Write-Log "========================================"
    Write-Log ""

    $PershingRows = $RawData | Where-Object {
        if ($null -ne $_._ProcessedBy) { return $false }
        $SourceCompany = $_."Source Company"
        if ([string]::IsNullOrWhiteSpace($SourceCompany)) { return $false }
        return ($SourceCompany -ilike "*$PershingKeyword*")
    }

    Write-Log "  Found $(@($PershingRows).Count) rows with Source Company containing '$PershingKeyword'"

    foreach ($Row in $PershingRows) { $Row._ProcessedBy = "Pershing" }
    Send-Count "Pershing (manual NetX check)" @($PershingRows).Count

    Write-Log "  (Rows marked for Raw DNU highlight - excluded from call lists, manual NetX check required)" -Level SUCCESS
    Write-Log ""


    # ====================================================================
    # STEP 11: CLASSIFY UNCLAIMED ROWS FOR ADVISOR & CONTRA FIRM LISTS
    # Old "Call Sheet" intermediate tab is gone - we go straight from
    # unclaimed rows to the two call list tabs.
    # ====================================================================

    Send-Status 3 48 "Preparing call lists"

    Write-Log "========================================"
    Write-Log "STEP 11: Classifying Unclaimed Rows for Call Lists"
    Write-Log "========================================"
    Write-Log ""

    $CallListRows = $RawData | Where-Object { $null -eq $_._ProcessedBy }

    Write-Log "  Found $($CallListRows.Count) unclaimed rows for Call Lists"

    foreach ($Row in $CallListRows) {
        $Row | Add-Member -NotePropertyName "Callback Date" -NotePropertyValue "" -Force
    }

    $CallListColumnOrder = @(
        "Source Company",
        "Client Name",
        "Notes",
        "Callback Date",
        "Source Account",
        "OES",
        "Rollover ID",
        "Representative",
        "LOA Date",
        "TPA Authorization Requested",
        "TPA Approval",
        "Expected Amount",
        "Received Amount",
        "Target Account",
        "Target Product Type",
        "Target Product",
        "Status",
        "Client Initiated",
        "Internal",
        "Funded at a Later Date"
    )

    if ($CallListRows.Count -gt 0) {
        $CallListActualProps = $CallListRows[0].PSObject.Properties.Name | Where-Object { $_ -notin @("_ProcessedBy","_IsRedCompany") }
    } else {
        $CallListActualProps = @()
    }
    $CallListOrderedProps = @()
    foreach ($Col in $CallListColumnOrder) {
        if ($Col -in $CallListActualProps) { $CallListOrderedProps += $Col }
    }
    foreach ($Col in $CallListActualProps) {
        if ($Col -notin $CallListOrderedProps) { $CallListOrderedProps += $Col }
    }


    # ====================================================================
    # STEP 12: SPLIT CALL LIST ROWS INTO ADVISOR (RED) AND CONTRA FIRM (BLUE)
    # (Pershing rows never reach this step - claimed in Step 10B.)
    # (DNC Rep marking removed - handled outside this script.)
    #
    # NEAR-MISS CAPTURE: any Blue row whose best score against a
    # redCompanies entry landed between $Script:NearMissReviewFloor and the
    # match threshold is recorded and printed in a REVIEW banner at the end
    # of the run. This is the safety net for a DNC firm whose Source Company
    # is worded differently enough to fall under the threshold. It will NOT
    # catch a firm missing from redCompanies entirely - that scores near
    # zero. Only a config/DNC-doc reconciliation catches those.
    # ====================================================================

    Send-Status 3 55 "Splitting Advisor and Contra Firm"

    Write-Log "========================================"
    Write-Log "STEP 12: Splitting Call List into Advisor (Red) and Contra Firm (Blue)"
    Write-Log "========================================"
    Write-Log ""

    if ($Script:DebugLoggingEnabled) {
        Write-Log "  Red Company config entries ($($RedCompanies.Count)):"
        foreach ($Comp in $RedCompanies) { Write-Log "    - '$Comp'" -Level DEBUG }
        Write-Log ""
    }

    $AdvisorRows    = @()
    $ContraFirmRows = @()
    $RedMatchSummary  = @{}
    $BlueCompanies    = @{}
    $NearMissReview   = @{}

    foreach ($Row in $CallListRows) {
        $SourceCompany = $Row."Source Company"

        $CompanyMatchResult = Test-FuzzyMatchAgainstList `
            -DataValue $SourceCompany `
            -ConfigValues $RedCompanies `
            -Label "RedCompany"

        $IsRed = $CompanyMatchResult.IsMatch
        $Row | Add-Member -NotePropertyName "_IsRedCompany" -NotePropertyValue $IsRed -Force

        if ($IsRed) {
            $AdvisorRows += $Row
            $MatchKey = $CompanyMatchResult.MatchedConfig
            if (-not $RedMatchSummary.ContainsKey($MatchKey)) {
                $RedMatchSummary[$MatchKey] = @{ Count = 0; Tier = $CompanyMatchResult.Tier; Examples = [System.Collections.Generic.List[string]]::new() }
            }
            $RedMatchSummary[$MatchKey].Count++
            if ($RedMatchSummary[$MatchKey].Examples.Count -lt 3) { $RedMatchSummary[$MatchKey].Examples.Add($SourceCompany) }
        }
        else {
            $ContraFirmRows += $Row
            if (-not [string]::IsNullOrWhiteSpace($SourceCompany)) {
                $BlueKey = $SourceCompany.Trim()
                if (-not $BlueCompanies.ContainsKey($BlueKey)) { $BlueCompanies[$BlueKey] = 0 }
                $BlueCompanies[$BlueKey]++

                if ($CompanyMatchResult.Score -ge $Script:NearMissReviewFloor -and
                    -not [string]::IsNullOrWhiteSpace($CompanyMatchResult.MatchedConfig)) {

                    # Only flag if the two names share a token that actually
                    # means something. Overlap on 'retirement' or 'funds'
                    # alone is noise and buries the real hits.
                    $SharedTokens = Get-SharedDistinctiveTokens `
                        -ValueA $BlueKey `
                        -ValueB $CompanyMatchResult.MatchedConfig

                    if ($SharedTokens.Count -gt 0) {
                        $NearMissKey = "$BlueKey  ~  $($CompanyMatchResult.MatchedConfig)"
                        if (-not $NearMissReview.ContainsKey($NearMissKey)) {
                            $NearMissReview[$NearMissKey] = @{
                                Count  = 0
                                Score  = $CompanyMatchResult.Score
                                Shared = ($SharedTokens -join ', ')
                            }
                        }
                        $NearMissReview[$NearMissKey].Count++
                    }
                }
            }
        }
    }

    Write-Log ""
    Write-Log "  === RED / BLUE MATCH SUMMARY ===" -Level SUCCESS
    Write-Log "  Total Red (Advisor) rows: $($AdvisorRows.Count)" -Level SUCCESS
    Write-Log "  Total Blue (Contra Firm) rows: $($ContraFirmRows.Count)" -Level SUCCESS
    Write-Log "  Pershing rows excluded (manual NetX check - see reminder at end): $(@($PershingRows).Count)" -Level SUCCESS
    Write-Log "  Near-miss companies flagged for review (see banner at end): $($NearMissReview.Count)" -Level SUCCESS

    Send-Count "Advisor Call List (red)" @($AdvisorRows).Count
    Send-Count "Contra Firm Call List (blue)" @($ContraFirmRows).Count

    if ($Script:DebugLoggingEnabled) {
        if ($RedMatchSummary.Count -gt 0) {
            Write-Log "  Matched config entries:" -Level DEBUG
            foreach ($Key in ($RedMatchSummary.Keys | Sort-Object)) {
                $Info = $RedMatchSummary[$Key]
                Write-Log "    Config: '$Key' matched $($Info.Count) rows (tier=$($Info.Tier)) e.g. '$($Info.Examples -join "', '")'" -Level DEBUG
            }
        }
        if ($BlueCompanies.Count -gt 0) {
            Write-Log "  Unique Blue (non-matched) companies ($($BlueCompanies.Count)):" -Level DEBUG
            foreach ($Key in ($BlueCompanies.Keys | Sort-Object)) {
                Write-Log "    '$Key' ($($BlueCompanies[$Key]) rows)" -Level DEBUG
            }
        }
    }


    # ====================================================================
    # STEP 13: MERGE OLD-REPORT NOTES & CALLBACK DATES INTO CALL LISTS
    # (Done in-memory before export so the exported sheets already contain
    # the carried-over notes/callback values.)
    # ====================================================================

    Send-Status 4 62 "Merging old report data"

    Write-Log "========================================"
    Write-Log "STEP 13: Merging old-report data into Call Lists"
    Write-Log "========================================"
    Write-Log ""

    # We'll still load old Advisor / Contra Firm Call Lists for full
    # formatting copy-over in Step 16.
    $OldAdvisorData    = @{}
    $OldContraFirmData = @{}

    if (-not [string]::IsNullOrWhiteSpace($OldOutstandingROReport) -and (Test-Path $OldOutstandingROReport)) {
        $OldSheetNames = (Get-ExcelSheetInfo -Path $OldOutstandingROReport).Name

        if ("Advisor Call List" -in $OldSheetNames) {
            Read-OldReportSheetData `
                -FilePath $OldOutstandingROReport `
                -SheetName "Advisor Call List" `
                -Into $OldAdvisorData `
                -IncludeCallbackDate `
                -IncludeSourceCompanyFormatting
            Write-Log "  Loaded $($OldAdvisorData.Count) entries from old Advisor Call List"
        }

        if ("Contra Firm Call List" -in $OldSheetNames) {
            Read-OldReportSheetData `
                -FilePath $OldOutstandingROReport `
                -SheetName "Contra Firm Call List" `
                -Into $OldContraFirmData `
                -IncludeCallbackDate `
                -IncludeSourceCompanyFormatting
            Write-Log "  Loaded $($OldContraFirmData.Count) entries from old Contra Firm Call List"
        }

        # Pre-fill in-memory Notes / Callback Date so they land in the exported sheets
        $MergeMatchAdvisor = 0
        foreach ($Row in $AdvisorRows) {
            $RID = $Row."Rollover ID"
            if ([string]::IsNullOrWhiteSpace($RID)) { continue }
            $RID = $RID.ToString().Trim()
            if ($OldAdvisorData.ContainsKey($RID)) {
                $Entry = $OldAdvisorData[$RID]
                if (-not [string]::IsNullOrWhiteSpace($Entry.Notes) -and [string]::IsNullOrWhiteSpace($Row.Notes)) {
                    $Row.Notes = $Entry.Notes
                }
                if (-not [string]::IsNullOrWhiteSpace($Entry.CallbackDate) -and [string]::IsNullOrWhiteSpace($Row."Callback Date")) {
                    $Row."Callback Date" = $Entry.CallbackDate
                }
                $MergeMatchAdvisor++
            }
        }

        $MergeMatchContra = 0
        foreach ($Row in $ContraFirmRows) {
            $RID = $Row."Rollover ID"
            if ([string]::IsNullOrWhiteSpace($RID)) { continue }
            $RID = $RID.ToString().Trim()
            if ($OldContraFirmData.ContainsKey($RID)) {
                $Entry = $OldContraFirmData[$RID]
                if (-not [string]::IsNullOrWhiteSpace($Entry.Notes) -and [string]::IsNullOrWhiteSpace($Row.Notes)) {
                    $Row.Notes = $Entry.Notes
                }
                if (-not [string]::IsNullOrWhiteSpace($Entry.CallbackDate) -and [string]::IsNullOrWhiteSpace($Row."Callback Date")) {
                    $Row."Callback Date" = $Entry.CallbackDate
                }
                $MergeMatchContra++
            }
        }

        Write-Log "  In-memory pre-merge: Advisor matched=$MergeMatchAdvisor, Contra Firm matched=$MergeMatchContra" -Level SUCCESS
    }
    else {
        Write-Log "  No old report provided - skipping in-memory pre-merge" -Level WARNING
    }

    Write-Log ""


    # ====================================================================
    # STEP 14: EXPORT ADVISOR & CONTRA FIRM CALL LISTS
    # ====================================================================

    Send-Status 5 70 "Exporting call lists"

    Write-Log "========================================"
    Write-Log "STEP 14: Exporting Advisor & Contra Firm Call Lists"
    Write-Log "========================================"
    Write-Log ""

    if ($AdvisorRows.Count -gt 0) {
        $AdvisorRows |
            Select-Object -Property $CallListOrderedProps |
            Export-Excel -Path $OutputFilePath `
                         -WorksheetName "Advisor Call List" `
                         -ClearSheet `
                         -AutoSize `
                         -AutoFilter `
                         -FreezeTopRow `
                         -BoldTopRow
        Write-Log "  Advisor Call List: $($AdvisorRows.Count) rows" -Level SUCCESS
    }
    else {
        Write-Log "  No Advisor Call List rows" -Level WARNING
    }

    if ($ContraFirmRows.Count -gt 0) {
        $ContraFirmRows |
            Select-Object -Property $CallListOrderedProps |
            Export-Excel -Path $OutputFilePath `
                         -WorksheetName "Contra Firm Call List" `
                         -ClearSheet `
                         -AutoSize `
                         -AutoFilter `
                         -FreezeTopRow `
                         -BoldTopRow
        Write-Log "  Contra Firm Call List: $($ContraFirmRows.Count) rows" -Level SUCCESS
    }
    else {
        Write-Log "  No Contra Firm Call List rows" -Level WARNING
    }

    Write-Log ""


    # ====================================================================
    # STEP 15: EXPORT RAW DNU SHEET
    # ====================================================================

    Send-Status 5 76 "Exporting Raw DNU"

    Write-Log "========================================"
    Write-Log "STEP 15: Exporting Raw DNU Sheet"
    Write-Log "========================================"
    Write-Log ""

    $RawData |
        Select-Object -Property $OrderedProperties |
        Export-Excel -Path $OutputFilePath `
                     -WorksheetName "Raw DNU" `
                     -ClearSheet `
                     -AutoSize `
                     -AutoFilter `
                     -FreezeTopRow `
                     -BoldTopRow

    Write-Log "  Raw DNU exported with $($RawData.Count) rows" -Level SUCCESS
    Send-Count "Raw DNU (all rows)" @($RawData).Count
    Write-Log ""


    # ====================================================================
    # STEP 16: COPY NOTES FROM OLD REPORT TO ALL SHEETS
    # ====================================================================

    Send-Status 6 82 "Copying notes from old report"

    Write-Log "========================================"
    Write-Log "STEP 16: Copying Notes (and Formatting) from Old Report"
    Write-Log "========================================"
    Write-Log ""

    if (-not [string]::IsNullOrWhiteSpace($OldOutstandingROReport) -and (Test-Path $OldOutstandingROReport)) {

        # Notes-only sheets: just Received now that Misc., State TRSs,
        # Pershing, Internals, and Pending LOA are no longer tabs.
        $NotesOnlySheets = @(
            "Received"
        )

        $CallListSheets = @(
            "Advisor Call List",
            "Contra Firm Call List"
        )

        $OutputPackage = Open-ExcelPackage -Path $OutputFilePath

        try {
            foreach ($SheetName in $NotesOnlySheets) {
                Write-Log "  Processing: $SheetName"

                $OldNotes = @{}
                Read-OldReportSheetData `
                    -FilePath $OldOutstandingROReport `
                    -SheetName $SheetName `
                    -Into $OldNotes

                $TargetWs = $OutputPackage.Workbook.Worksheets[$SheetName]
                if ($null -eq $TargetWs) {
                    Write-Log "    Sheet '$SheetName' not in output workbook - skipping" -Level WARNING
                    continue
                }

                Apply-OldDataToSheet `
                    -Worksheet $TargetWs `
                    -OldData $OldNotes `
                    -SheetName $SheetName
            }

            # Advisor / Contra Firm: full data + formatting copy
            foreach ($SheetName in $CallListSheets) {
                Write-Log "  Processing: $SheetName (with Callback Date + Source Company formatting)"

                $OldCallData = if ($SheetName -eq "Advisor Call List") { $OldAdvisorData } else { $OldContraFirmData }

                $TargetWs = $OutputPackage.Workbook.Worksheets[$SheetName]
                if ($null -eq $TargetWs) {
                    Write-Log "    Sheet '$SheetName' not in output workbook - skipping" -Level WARNING
                    continue
                }

                Apply-OldDataToSheet `
                    -Worksheet $TargetWs `
                    -OldData $OldCallData `
                    -SheetName $SheetName `
                    -IncludeCallbackDate `
                    -IncludeSourceCompanyFormatting
            }

            Close-ExcelPackage $OutputPackage
        }
        catch {
            Write-Log "Error during notes copy: $($_.Exception.Message)" -Level ERROR
            try { $OutputPackage.Dispose() } catch { }
            throw
        }

        Write-Log "  Notes copy complete" -Level SUCCESS
    }
    else {
        Write-Log "  No old report provided - skipping notes copy" -Level WARNING
    }

    Write-Log ""


    # ====================================================================
    # STEP 16B: CARRY OVER NIGO ROWS FROM OLD REPORT
    # Scan old Advisor / Contra Firm Call Lists for rows whose Callback
    # Date contains "NIGO". Any whose Rollover ID does NOT appear in the
    # new raw data are appended to the matching new sheet.
    # ====================================================================

    Send-Status 6 86 "Carrying over NIGO rows"

    Write-Log "========================================"
    Write-Log "STEP 16B: Carrying Over NIGO Rows from Old Report"
    Write-Log "========================================"
    Write-Log ""

    if (-not [string]::IsNullOrWhiteSpace($OldOutstandingROReport) -and (Test-Path $OldOutstandingROReport)) {

        $NewRolloverIDs = [System.Collections.Generic.HashSet[string]]::new()
        foreach ($Row in $RawData) {
            $RID = $Row."Rollover ID"
            if (-not [string]::IsNullOrWhiteSpace($RID)) {
                [void]$NewRolloverIDs.Add($RID.ToString().Trim())
            }
        }

        $NIGOSourceSheets = @("Advisor Call List", "Contra Firm Call List")

        $NigoPkg = Open-ExcelPackage -Path $OutputFilePath

        try {
            foreach ($SourceSheetName in $NIGOSourceSheets) {
                if ($Script:DebugLoggingEnabled) { Write-Log "  Scanning old '$SourceSheetName' for NIGO rows..." }

                $OldSheetNames = (Get-ExcelSheetInfo -Path $OldOutstandingROReport).Name
                if ($SourceSheetName -notin $OldSheetNames) {
                    Write-Log "    '$SourceSheetName' not found in old report - skipping" -Level WARNING
                    continue
                }

                $OldPkg = Open-ExcelPackage -Path $OldOutstandingROReport
                $NIGORows = [System.Collections.Generic.List[object]]::new()

                try {
                    $OldWs = $OldPkg.Workbook.Worksheets[$SourceSheetName]
                    if ($null -eq $OldWs -or $null -eq $OldWs.Dimension) {
                        Write-Log "    Sheet '$SourceSheetName' empty - skipping" -Level WARNING
                        continue
                    }

                    $OldStartRow = $OldWs.Dimension.Start.Row
                    $OldEndRow   = $OldWs.Dimension.End.Row
                    $OldEndCol   = $OldWs.Dimension.End.Column

                    $OldHeaderMap = @{}
                    for ($col = 1; $col -le $OldEndCol; $col++) {
                        $H = $OldWs.Cells[$OldStartRow, $col].Text.Trim()
                        if (-not [string]::IsNullOrWhiteSpace($H)) { $OldHeaderMap[$H.ToLower()] = $col }
                    }

                    $OldRIDCol      = $OldHeaderMap["rollover id"]
                    $OldCallbackCol = $OldHeaderMap["callback date"]

                    if (-not $OldRIDCol -or -not $OldCallbackCol) {
                        Write-Log "    Required columns missing in old '$SourceSheetName' - skipping" -Level WARNING
                        continue
                    }

                    for ($row = ($OldStartRow + 1); $row -le $OldEndRow; $row++) {
                        $RID          = $OldWs.Cells[$row, $OldRIDCol].Text.Trim()
                        $CallbackText = $OldWs.Cells[$row, $OldCallbackCol].Text.Trim()

                        if ($CallbackText -ilike "*NIGO*" -and
                            -not [string]::IsNullOrWhiteSpace($RID) -and
                            -not $NewRolloverIDs.Contains($RID)) {

                            $RowData = [ordered]@{}
                            foreach ($Key in $OldHeaderMap.Keys) {
                                $RowData[$Key] = $OldWs.Cells[$row, $OldHeaderMap[$Key]].Text
                            }
                            $NIGORows.Add($RowData)
                        }
                    }
                }
                finally {
                    $OldPkg.Dispose()
                }

                if ($Script:DebugLoggingEnabled) {
                    Write-Log "  Found $($NIGORows.Count) NIGO rows to carry over into '$SourceSheetName'" -Level SUCCESS
                }

                if ($NIGORows.Count -eq 0) { continue }

                $TargetWs = $NigoPkg.Workbook.Worksheets[$SourceSheetName]
                if ($null -eq $TargetWs) {
                    Write-Log "    '$SourceSheetName' not found in output workbook - skipping" -Level WARNING
                    continue
                }

                $TgtEndRow = $TargetWs.Dimension.End.Row
                $TgtEndCol = $TargetWs.Dimension.End.Column

                $TgtHeaderMap = @{}
                for ($col = 1; $col -le $TgtEndCol; $col++) {
                    $H = $TargetWs.Cells[1, $col].Text.Trim()
                    if (-not [string]::IsNullOrWhiteSpace($H)) { $TgtHeaderMap[$H.ToLower()] = $col }
                }

                $AppendRow = $TgtEndRow + 1
                foreach ($NIGORow in $NIGORows) {
                    foreach ($ColName in $NIGORow.Keys) {
                        if ($TgtHeaderMap.ContainsKey($ColName)) {
                            $TargetWs.Cells[$AppendRow, $TgtHeaderMap[$ColName]].Value = $NIGORow[$ColName]
                        }
                    }
                    $AppendRow++
                }

                if ($Script:DebugLoggingEnabled) {
                    Write-Log "  Appended $($NIGORows.Count) NIGO rows to '$SourceSheetName'" -Level SUCCESS
                }
            }

            Close-ExcelPackage $NigoPkg
        }
        catch {
            Write-Log "Error during NIGO carry-over: $($_.Exception.Message)" -Level ERROR
            try { $NigoPkg.Dispose() } catch { }
            throw
        }
    }
    else {
        Write-Log "  No old report provided - skipping NIGO carry-over" -Level WARNING
    }

    Write-Log ""


    # ====================================================================
    # STEP 17: APPLY ALL FORMATTING (COLORS, TAB COLORS, BORDERS, WIDTHS)
    # ====================================================================

    Send-Status 6 90 "Applying formatting"

    Write-Log "========================================"
    Write-Log "STEP 17: Applying Formatting"
    Write-Log "========================================"
    Write-Log ""

    $Pkg = Open-ExcelPackage -Path $OutputFilePath

    try {

        $HeaderBgColor     = Convert-HexToColor "#e97132"
        $HeaderFontColor   = [System.Drawing.Color]::Black

        $ReceivedColorObj        = Convert-HexToColor $ReceivedColor
        $StateTRSColorObj        = Convert-HexToColor $StateTRSsColor
        $TradingPlatformColorObj = Convert-HexToColor $TradingPlatformColor
        $PershingColorObj        = Convert-HexToColor $PershingColor
        $InternalsColorObj      = Convert-HexToColor $InternalsColor
        $FundingColorObj        = Convert-HexToColor $FundingReviewColor
        $PendingLOAColorObj     = Convert-HexToColor $PendingLOAColor
        $RedColorObj            = Convert-HexToColor $CallSheetRedColor
        $BlueColorObj           = Convert-HexToColor $CallSheetBlueColor

        # Used for Raw DNU row highlighting based on _ProcessedBy
        $ProcessedByColorMap = @{
            "Received"         = $ReceivedColorObj
            "State TRSs"       = $StateTRSColorObj
            "Trading Platform" = $TradingPlatformColorObj
            "Pershing"         = $PershingColorObj
            "Internals"        = $InternalsColorObj
            "Funding Review"   = $FundingColorObj
            "Pending LOA"      = $PendingLOAColorObj
        }

        # Tab colors for exported sheets only
        $TabColorMap = @{
            "Received"              = $ReceivedColorObj
            "Funding Review"        = $FundingColorObj
            "Advisor Call List"     = $RedColorObj
            "Contra Firm Call List" = $BlueColorObj
        }


        # ----------------------------------------------------------------
        # ITERATE ALL SHEETS: Apply common formatting
        # ----------------------------------------------------------------

        foreach ($Ws in $Pkg.Workbook.Worksheets) {
            $SheetName = $Ws.Name
            Write-Log "  Formatting: $SheetName"

            if ($null -eq $Ws.Dimension) {
                Write-Log "    Empty sheet - skipping formatting"
                continue
            }

            $StartRow = $Ws.Dimension.Start.Row
            $EndRow   = $Ws.Dimension.End.Row
            $StartCol = $Ws.Dimension.Start.Column
            $EndCol   = $Ws.Dimension.End.Column

            for ($col = $StartCol; $col -le $EndCol; $col++) {
                $HeaderCell = $Ws.Cells[$StartRow, $col]
                $HeaderCell.Style.Font.Bold = $true
                $HeaderCell.Style.Font.Color.SetColor($HeaderFontColor)
                $HeaderCell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                $HeaderCell.Style.Fill.BackgroundColor.SetColor($HeaderBgColor)
            }

            $AllCells = $Ws.Cells[$StartRow, $StartCol, $EndRow, $EndCol]
            $AllCells.Style.Border.Top.Style    = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
            $AllCells.Style.Border.Bottom.Style = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
            $AllCells.Style.Border.Left.Style   = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin
            $AllCells.Style.Border.Right.Style  = [OfficeOpenXml.Style.ExcelBorderStyle]::Thin

            if ($EndRow -ge 2) {
                $DataCells = $Ws.Cells[2, $StartCol, $EndRow, $EndCol]
                $DataCells.Style.HorizontalAlignment = [OfficeOpenXml.Style.ExcelHorizontalAlignment]::Left
            }

            for ($col = $StartCol; $col -le $EndCol; $col++) {
                $ColHeader = $Ws.Cells[$StartRow, $col].Text.Trim().ToLower()
                if ($ColHeader -eq "notes") { $Ws.Column($col).Width = 50 }
            }

            $Ws.View.FreezePanes(2, 1)

            if ($TabColorMap.ContainsKey($SheetName)) { $Ws.TabColor = $TabColorMap[$SheetName] }
        }


        # ----------------------------------------------------------------
        # RAW DNU: Apply row highlight colors based on _ProcessedBy
        # ----------------------------------------------------------------

        Write-Log "  Applying highlight colors to Raw DNU..."

        $RawDNUSheet = $Pkg.Workbook.Worksheets["Raw DNU"]

        if ($null -ne $RawDNUSheet -and $null -ne $RawDNUSheet.Dimension) {
            $RawEndRow = $RawDNUSheet.Dimension.End.Row
            $RawEndCol = $RawDNUSheet.Dimension.End.Column

            $ProcessedByLookup = @{}
            foreach ($Row in $RawData) {
                $RID = ""
                if ($Row."Rollover ID") { $RID = $Row."Rollover ID".ToString().Trim() }
                if (-not [string]::IsNullOrWhiteSpace($RID) -and $Row._ProcessedBy) {
                    $ProcessedByLookup[$RID] = $Row._ProcessedBy
                }
            }

            $RawRolloverIDCol = $null
            for ($col = 1; $col -le $RawEndCol; $col++) {
                if ($RawDNUSheet.Cells[1, $col].Text.Trim().ToLower() -eq "rollover id") {
                    $RawRolloverIDCol = $col
                    break
                }
            }

            if ($RawRolloverIDCol) {
                $ColoredCount = 0
                for ($row = 2; $row -le $RawEndRow; $row++) {
                    $RID = $RawDNUSheet.Cells[$row, $RawRolloverIDCol].Text.Trim()

                    if ($ProcessedByLookup.ContainsKey($RID)) {
                        $SheetClaim = $ProcessedByLookup[$RID]

                        if ($ProcessedByColorMap.ContainsKey($SheetClaim)) {
                            $HighlightColor = $ProcessedByColorMap[$SheetClaim]
                            for ($col = 1; $col -le $RawEndCol; $col++) {
                                $RawDNUSheet.Cells[$row, $col].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                                $RawDNUSheet.Cells[$row, $col].Style.Fill.BackgroundColor.SetColor($HighlightColor)
                            }
                            $ColoredCount++
                        }
                    }
                }
                Write-Log "    Highlighted $ColoredCount rows in Raw DNU" -Level SUCCESS
            }
            else {
                Write-Log "    Could not find 'Rollover ID' column in Raw DNU sheet" -Level WARNING
            }
        }


        # ----------------------------------------------------------------
        # ADVISOR CALL LIST: Color all Source Company cells red
        # ----------------------------------------------------------------

        Write-Log "  Applying coloring to Advisor Call List..."

        $AdvisorWs = $Pkg.Workbook.Worksheets["Advisor Call List"]

        if ($null -ne $AdvisorWs -and $null -ne $AdvisorWs.Dimension) {
            $AEndRow = $AdvisorWs.Dimension.End.Row
            $AEndCol = $AdvisorWs.Dimension.End.Column

            $ASrcCompCol = $null
            for ($col = 1; $col -le $AEndCol; $col++) {
                if ($AdvisorWs.Cells[1, $col].Text.Trim().ToLower() -eq "source company") {
                    $ASrcCompCol = $col
                    break
                }
            }

            $RedCount = 0

            if ($ASrcCompCol) {
                for ($row = 2; $row -le $AEndRow; $row++) {
                    $Cell = $AdvisorWs.Cells[$row, $ASrcCompCol]
                    $Cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $Cell.Style.Fill.BackgroundColor.SetColor($RedColorObj)
                    $RedCount++
                }
                Write-Log "    Red cells: $RedCount" -Level SUCCESS
            }
            else {
                Write-Log "    'Source Company' column not found on Advisor Call List" -Level WARNING
            }
        }


        # ----------------------------------------------------------------
        # CONTRA FIRM CALL LIST: Color all Source Company cells blue
        # ----------------------------------------------------------------

        Write-Log "  Applying coloring to Contra Firm Call List..."

        $ContraWs = $Pkg.Workbook.Worksheets["Contra Firm Call List"]

        if ($null -ne $ContraWs -and $null -ne $ContraWs.Dimension) {
            $CEndRow = $ContraWs.Dimension.End.Row
            $CEndCol = $ContraWs.Dimension.End.Column

            $CSrcCompCol = $null
            for ($col = 1; $col -le $CEndCol; $col++) {
                if ($ContraWs.Cells[1, $col].Text.Trim().ToLower() -eq "source company") {
                    $CSrcCompCol = $col
                    break
                }
            }

            $BlueCount = 0

            if ($CSrcCompCol) {
                for ($row = 2; $row -le $CEndRow; $row++) {
                    $Cell = $ContraWs.Cells[$row, $CSrcCompCol]
                    $Cell.Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
                    $Cell.Style.Fill.BackgroundColor.SetColor($BlueColorObj)
                    $BlueCount++
                }
                Write-Log "    Blue cells: $BlueCount" -Level SUCCESS
            }
            else {
                Write-Log "    'Source Company' column not found on Contra Firm Call List" -Level WARNING
            }
        }


        # ----------------------------------------------------------------
        # REORDER WORKSHEET TABS
        # ----------------------------------------------------------------

        Write-Log "  Reordering worksheet tabs..."

        $DesiredTabOrder = @(
            "Raw DNU",
            "Advisor Call List",
            "Contra Firm Call List",
            "Funding Review",
            "Received"
        )

        $ExistingSheetNames = $Pkg.Workbook.Worksheets | ForEach-Object { $_.Name }

        $OrderedExisting = @()
        foreach ($Name in $DesiredTabOrder) {
            if ($Name -in $ExistingSheetNames) { $OrderedExisting += $Name }
        }
        foreach ($Name in $ExistingSheetNames) {
            if ($Name -notin $OrderedExisting) { $OrderedExisting += $Name }
        }

        for ($i = 0; $i -lt $OrderedExisting.Count; $i++) {
            $SheetName = $OrderedExisting[$i]
            $Sheet = $Pkg.Workbook.Worksheets[$SheetName]
            if ($Sheet) { $Pkg.Workbook.Worksheets.MoveToEnd($SheetName) }
        }

        for ($i = 1; $i -lt $OrderedExisting.Count; $i++) {
            $CurrentSheet  = $Pkg.Workbook.Worksheets[$OrderedExisting[$i]]
            $PreviousSheet = $Pkg.Workbook.Worksheets[$OrderedExisting[$i - 1]]
            if ($CurrentSheet -and $PreviousSheet) {
                $Pkg.Workbook.Worksheets.MoveAfter($OrderedExisting[$i], $OrderedExisting[$i - 1])
            }
        }

        Write-Log "    Tab order: $($OrderedExisting -join ', ')" -Level SUCCESS


        # ----------------------------------------------------------------
        # CLEANUP: Remove leftover default sheets
        # ----------------------------------------------------------------

        try {
            $DefaultSheet = $Pkg.Workbook.Worksheets["Sheet1"]
            if ($DefaultSheet -and $Pkg.Workbook.Worksheets.Count -gt 1) {
                $Pkg.Workbook.Worksheets.Delete("Sheet1")
                Write-Log "    Removed leftover 'Sheet1'"
            }
        } catch { }

        try {
            $OriginalSheetName = (Get-ExcelSheetInfo -Path $RawInputFile)[0].Name
            if ($OriginalSheetName -ne "Raw DNU") {
                $OrigSheet = $Pkg.Workbook.Worksheets[$OriginalSheetName]
                if ($OrigSheet -and $Pkg.Workbook.Worksheets.Count -gt 1) {
                    $Pkg.Workbook.Worksheets.Delete($OriginalSheetName)
                    Write-Log "    Removed original sheet '$OriginalSheetName'"
                }
            }
        } catch { }


        # ----------------------------------------------------------------
        # SAVE
        # ----------------------------------------------------------------
        Close-ExcelPackage $Pkg

        Write-Log ""
        Write-Log "  All formatting applied and workbook saved" -Level SUCCESS

    }
    catch {
        Write-Log "Error during formatting: $($_.Exception.Message)" -Level ERROR
        Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level ERROR
        try { $Pkg.Dispose() } catch { }
        throw
    }

    Write-Log ""


    # ====================================================================
    # STEP 18: AUTO-FILL CALLBACK DATES FROM LOA DATE + 14 DAYS
    # For each row in Advisor / Contra Firm Call List, if Callback Date is
    # still empty, set it to LOA Date + 14 days. Rows that already have a
    # Callback Date (from old-report merge or NIGO carry-over) are left
    # untouched.
    # ====================================================================

    Send-Status 6 96 "Auto-filling callback dates"

    Write-Log "========================================"
    Write-Log "STEP 18: Auto-Fill Callback Dates (LOA Date + 14 Days)"
    Write-Log "========================================"
    Write-Log ""

    $CallbackSheets = @("Advisor Call List", "Contra Firm Call List")

    $CbPkg = Open-ExcelPackage -Path $OutputFilePath

    try {
        foreach ($SheetName in $CallbackSheets) {
            Write-Log "  Processing: $SheetName"

            $CbWs = $CbPkg.Workbook.Worksheets[$SheetName]
            if ($null -eq $CbWs -or $null -eq $CbWs.Dimension) {
                Write-Log "    Sheet '$SheetName' not found or empty - skipping" -Level WARNING
                continue
            }

            $CbEndRow = $CbWs.Dimension.End.Row
            $CbEndCol = $CbWs.Dimension.End.Column

            $CbHeaderMap = @{}
            for ($col = 1; $col -le $CbEndCol; $col++) {
                $H = $CbWs.Cells[1, $col].Text.Trim()
                if (-not [string]::IsNullOrWhiteSpace($H)) {
                    $CbHeaderMap[$H.ToLower()] = $col
                }
            }

            $LOADateCol  = $CbHeaderMap["loa date"]
            $CallbackCol = $CbHeaderMap["callback date"]

            if (-not $LOADateCol) {
                Write-Log "    'LOA Date' column not found in '$SheetName' - skipping" -Level WARNING
                continue
            }
            if (-not $CallbackCol) {
                Write-Log "    'Callback Date' column not found in '$SheetName' - skipping" -Level WARNING
                continue
            }

            $FilledCount  = 0
            $SkippedCount = 0
            $NoLOACount   = 0

            for ($row = 2; $row -le $CbEndRow; $row++) {
                $ExistingCallback = $CbWs.Cells[$row, $CallbackCol].Text.Trim()

                if (-not [string]::IsNullOrWhiteSpace($ExistingCallback)) {
                    $SkippedCount++
                    continue
                }

                $LOACell     = $CbWs.Cells[$row, $LOADateCol]
                $LOARawValue = $LOACell.Value
                $LOAText     = $LOACell.Text.Trim()

                $ParsedLOA = $null

                if ($LOARawValue -is [DateTime]) {
                    $ParsedLOA = $LOARawValue
                }
                elseif (-not [string]::IsNullOrWhiteSpace($LOAText)) {
                    $TempDate = [DateTime]::MinValue
                    if ([DateTime]::TryParse($LOAText, [ref]$TempDate)) {
                        $ParsedLOA = $TempDate
                    }
                }

                if ($null -eq $ParsedLOA) {
                    $NoLOACount++
                    continue
                }

                $CallbackDate = $ParsedLOA.AddDays(14)
                $CbWs.Cells[$row, $CallbackCol].Value = $CallbackDate
                $CbWs.Cells[$row, $CallbackCol].Style.Numberformat.Format = "mm/dd/yyyy"
                $FilledCount++
            }

            Write-Log "    Filled: $FilledCount | Already set: $SkippedCount | No LOA Date: $NoLOACount" -Level SUCCESS
        }

        Close-ExcelPackage $CbPkg
        Write-Log ""
        Write-Log "  Callback date auto-fill complete" -Level SUCCESS

    }
    catch {
        Write-Log "Error during callback date auto-fill: $($_.Exception.Message)" -Level ERROR
        try { $CbPkg.Dispose() } catch { }
        throw
    }

    Write-Log ""


    # ====================================================================
    # COMPLETION
    # ====================================================================

    Write-Log "========================================"
    Write-Log "Processing Complete"
    Write-Log "========================================"
    Write-Log "Output file: $OutputFilePath" -Level SUCCESS
    Write-Log "Log file: $LogFile" -Level SUCCESS


    # ====================================================================
    # MANUAL STEP REMINDER: PERSHING NETX CHECK
    # ====================================================================

    Write-Log ""
    Write-Log "****************************************************************" -Level WARNING
    Write-Log "*  MANUAL STEP REQUIRED: PERSHING NETX CHECK                   *" -Level WARNING
    Write-Log "****************************************************************" -Level WARNING

    $PershingRowsArr = @($PershingRows)

    if ($PershingRowsArr.Count -eq 0) {
        Write-Log "  No Pershing rows were excluded this run - nothing to check." -Level WARNING
    }
    else {
        Write-Log "  $($PershingRowsArr.Count) Pershing row(s) were excluded from the call lists." -Level WARNING
        Write-Log "  These are highlighted on the Raw DNU tab. For EACH one, check" -Level WARNING
        Write-Log "  whether the LOA was uploaded into NetX, then manually move it:" -Level WARNING
        Write-Log "    - LOA in NetX      -> Contra Firm Call List" -Level WARNING
        Write-Log "    - LOA NOT in NetX  -> Advisor Call List" -Level WARNING
        Write-Log "" -Level WARNING
        Write-Log "  Rows to check:" -Level WARNING

        # Source Account leads the line: it is what you actually search NetX on,
        # so it should be the first thing your eye lands on rather than buried
        # behind the OES and rollover IDs.
        foreach ($PRow in $PershingRowsArr) {
            $PSrcAcct = ""
            $PClient  = ""
            $POES     = ""
            $PRID     = ""
            if (-not [string]::IsNullOrWhiteSpace($PRow."Source Account")) { $PSrcAcct = $PRow."Source Account".ToString().Trim() }
            if (-not [string]::IsNullOrWhiteSpace($PRow."Client Name"))    { $PClient  = $PRow."Client Name".ToString().Trim() }
            if (-not [string]::IsNullOrWhiteSpace($PRow."OES"))            { $POES     = $PRow."OES".ToString().Trim() }
            if (-not [string]::IsNullOrWhiteSpace($PRow."Rollover ID"))    { $PRID     = $PRow."Rollover ID".ToString().Trim() }
            Write-Log "    Source Account: $PSrcAcct | OES: $POES | Rollover ID: $PRID | Client: $PClient" -Level WARNING
            Send-Marker "@@PERSH Source Account: $PSrcAcct | OES: $POES | Rollover ID: $PRID | Client: $PClient"
        }
    }

    Write-Log "****************************************************************" -Level WARNING


    # ====================================================================
    # REVIEW: BLUE COMPANIES THAT NEARLY MATCHED A RED ENTRY
    # These landed on the Contra Firm Call List but scored close enough to
    # a redCompanies entry to be worth a human look. A DNC firm whose
    # Source Company is worded differently than the config entry shows up
    # here instead of quietly staying blue.
    # ====================================================================

    if ($NearMissReview.Count -gt 0) {
        Write-Log ""
        Write-Log "****************************************************************" -Level WARNING
        Write-Log "*  REVIEW: BLUE COMPANIES THAT NEARLY MATCHED A RED ENTRY      *" -Level WARNING
        Write-Log "****************************************************************" -Level WARNING
        Write-Log "  $($NearMissReview.Count) company name(s) scored at or above $($Script:NearMissReviewFloor)" -Level WARNING
        Write-Log "  against a redCompanies entry but below the $($Script:FuzzyMatchThreshold) match" -Level WARNING
        Write-Log "  threshold, so they stayed on the Contra Firm Call List." -Level WARNING
        Write-Log "" -Level WARNING
        Write-Log "  If any of these IS the firm it resembles, add its exact Source" -Level WARNING
        Write-Log "  Company wording to redCompanies in config.json (and update the" -Level WARNING
        Write-Log "  Company Information List + the Do Not Call resource to match)." -Level WARNING
        Write-Log "" -Level WARNING
        Write-Log "  Source Company  ~  closest redCompanies entry:" -Level WARNING

        foreach ($Key in ($NearMissReview.Keys | Sort-Object)) {
            $NM = $NearMissReview[$Key]
            Write-Log "    $Key   (score $([math]::Round($NM.Score, 2)), $($NM.Count) rows, shared: $($NM.Shared))" -Level WARNING
            Send-Marker "@@NEARMISS $Key   (score $([math]::Round($NM.Score, 2)), $($NM.Count) rows, shared: $($NM.Shared))"
        }

        Write-Log "****************************************************************" -Level WARNING
    }

    # ====================================================================
    # COMPLETION MARKER
    # Emitted last, after the reminder banners, so the GUI already has every
    # @@COUNT / @@PERSH / @@NEARMISS line before the status flips to Done.
    # ====================================================================

    Send-Marker "@@DONE workbook=$OutputFilePath folder=$OutputFolder log=$LogFile"

}
catch {
    Write-Log "Fatal error: $($_.Exception.Message)" -Level ERROR
    Write-Log "Stack trace: $($_.ScriptStackTrace)" -Level ERROR
    Send-Fail $_.Exception.Message
}

if (-not $isJobMode) {
    Write-Host ""
    Write-Host "Output folder: $OutputFolder" -ForegroundColor Cyan
    Write-Host ""
}
Wait-ForOperator
