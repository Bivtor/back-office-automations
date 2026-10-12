# normalize_dollars.ps1
# Strips $ and , from the dollar amount at the end of remit filenames.
# Splits the filename (minus extension) on spaces and only inspects the
# LAST token -- so business names, parentheses, dates, etc. are physically
# untouchable.

Write-Host "=== Remit Filename Normalizer ===" -ForegroundColor Green
Write-Host ""

# --- Guess today's ACH folder ---
$Today    = Get-Date
$Year     = $Today.Year
$Month    = $Today.ToString("MM")
$Day      = $Today.Day
$ACHBase  = "Q:\CORE\PURCHASE TRADING\ACH Remittance Detail"

# Try leading-zero day (01) first, then no-zero (1). If neither exists yet,
# suggest the leading-zero form as the guess.
$ACHMonthFolder = "$ACHBase\ACH-$Year\$Month - ACH"
$ACHDayZero     = "$ACHMonthFolder\$($Day.ToString('00'))"   # 01 form
$ACHDayNo       = "$ACHMonthFolder\$Day"                       # 1 form

if (Test-Path $ACHDayZero) {
    $Guessed = $ACHDayZero
} elseif (Test-Path $ACHDayNo) {
    $Guessed = $ACHDayNo
} else {
    $Guessed = $ACHDayZero
}

Write-Host "Today's ACH folder?" -ForegroundColor Yellow
Write-Host "Guessed: $Guessed" -ForegroundColor Cyan
Write-Host "Press Enter to use the guess, or paste a different path:"
$Input = Read-Host "ACH folder"

if ([string]::IsNullOrWhiteSpace($Input)) {
    $TargetDir = $Guessed
} else {
    $TargetDir = $Input.Trim().Trim('"', "'")
}

if (-not (Test-Path $TargetDir)) {
    Write-Host ""
    Write-Host "ERROR: folder does not exist: $TargetDir" -ForegroundColor Red
    Read-Host "Press Enter to exit"
    exit 1
}

Write-Host ""
Write-Host "Scanning: $TargetDir" -ForegroundColor Cyan
Write-Host ""

# --- Scan ---
$files    = Get-ChildItem -LiteralPath $TargetDir -File
$changes  = @()
$problems = @()

foreach ($f in $files) {
    $name     = $f.Name
    $base     = [System.IO.Path]::GetFileNameWithoutExtension($name)
    $ext      = [System.IO.Path]::GetExtension($name)

    # Split base on spaces; grab the last token.
    $tokens   = $base -split ' '
    if ($tokens.Count -eq 0) { continue }

    # Determine which index holds the dollar amount.
    # Normally it's the last token, but if the last token looks like "(N)"
    # (a duplicate marker like "(2)"), the amount is the token before it.
    $targetIdx = $tokens.Count - 1
    if ($tokens[$targetIdx] -match '^\(\d+\)$') {
        $targetIdx = $targetIdx - 1
        if ($targetIdx -lt 0) { continue }
    }
    $target = $tokens[$targetIdx]

    # Only normalize if the target token contains $ or , -- otherwise leave alone.
    if ($target -notmatch '[\$,]') { continue }

    # Strip $ and , from the target token only.
    $cleanTarget = $target -replace '[\$,]', ''

    # Sanity: cleaned token should look like a number (digits + optional decimal).
    # If it doesn't, skip rather than rename to something weird.
    if ($cleanTarget -notmatch '^\d+(\.\d+)?$') {
        $problems += [pscustomobject]@{
            File   = $name
            Reason = "target token '$target' didn't look like a dollar amount after cleanup ('$cleanTarget') -- skipped"
        }
        continue
    }

    # Rebuild filename: replace target token, re-add extension.
    $tokens[$targetIdx] = $cleanTarget
    $newBase            = $tokens -join ' '
    $newName            = "$newBase$ext"

    if ($newName -eq $name) { continue }

    # Collision check.
    $newPath = Join-Path $f.DirectoryName $newName
    if ((Test-Path -LiteralPath $newPath) -and ($newPath -ne $f.FullName)) {
        $problems += [pscustomobject]@{
            File   = $name
            Reason = "target name '$newName' already exists -- skipped"
        }
        continue
    }

    $changes += [pscustomobject]@{
        Old         = $name
        New         = $newName
        FullPath    = $f.FullName
        DirtyAmount = $target
        CleanAmount = $cleanTarget
    }
}

# --- Report ---
Write-Host "Files scanned:   $($files.Count)" -ForegroundColor White
Write-Host "Files to change: $($changes.Count)" -ForegroundColor Green
if ($problems.Count -gt 0) {
    Write-Host "Problems:        $($problems.Count)" -ForegroundColor Red
}
Write-Host ""

if ($changes.Count -eq 0 -and $problems.Count -eq 0) {
    Write-Host "Nothing to do." -ForegroundColor Gray
    Write-Host ""
    Read-Host "Press Enter to exit"
    exit 0
}

if ($changes.Count -gt 0) {
    Write-Host "Proposed renames:" -ForegroundColor Yellow
    foreach ($c in $changes) {
        # Split each name into: prefix | amount | suffix, so we can color
        # only the amount piece. The amount may be followed by " (N)" and
        # then the extension, so we locate the amount by its last index
        # rather than assuming it's at the end.
        $oldDirtyIdx = $c.Old.LastIndexOf($c.DirtyAmount)
        $newCleanIdx = $c.New.LastIndexOf($c.CleanAmount)
        $oldPrefix   = $c.Old.Substring(0, $oldDirtyIdx)
        $oldSuffix   = $c.Old.Substring($oldDirtyIdx + $c.DirtyAmount.Length)
        $newPrefix   = $c.New.Substring(0, $newCleanIdx)
        $newSuffix   = $c.New.Substring($newCleanIdx + $c.CleanAmount.Length)

        Write-Host "  " -NoNewline
        Write-Host $oldPrefix -ForegroundColor Gray -NoNewline
        Write-Host $c.DirtyAmount -ForegroundColor Red -NoNewline
        Write-Host $oldSuffix -ForegroundColor Gray -NoNewline
        Write-Host "  ->  " -ForegroundColor DarkGray -NoNewline
        Write-Host $newPrefix -ForegroundColor Gray -NoNewline
        Write-Host $c.CleanAmount -ForegroundColor Green -NoNewline
        Write-Host $newSuffix -ForegroundColor Gray
    }
    Write-Host ""
}

if ($problems.Count -gt 0) {
    Write-Host "Problems (will NOT be renamed):" -ForegroundColor Red
    foreach ($p in $problems) {
        Write-Host "  $($p.File)" -ForegroundColor Red
        Write-Host "     $($p.Reason)" -ForegroundColor DarkRed
    }
    Write-Host ""
}

if ($changes.Count -eq 0) {
    Read-Host "Press Enter to exit"
    exit 0
}

# --- Confirm + apply ---
Write-Host "Press Enter to apply these renames, or Ctrl+C to cancel." -ForegroundColor Yellow
Read-Host | Out-Null

Write-Host ""
Write-Host "Applying..." -ForegroundColor Magenta
$renamed = 0
$failed  = @()
foreach ($c in $changes) {
    try {
        Rename-Item -LiteralPath $c.FullPath -NewName $c.New -ErrorAction Stop
        $renamed++
    } catch {
        $failed += [pscustomobject]@{ File = $c.Old; Error = $_.Exception.Message }
    }
}

Write-Host ""
Write-Host "Renamed: $renamed" -ForegroundColor Green
if ($failed.Count -gt 0) {
    Write-Host "Failed:  $($failed.Count)" -ForegroundColor Red
    foreach ($x in $failed) {
        Write-Host "  $($x.File): $($x.Error)" -ForegroundColor Red
    }
}

Write-Host ""
Read-Host "Press Enter to exit"
