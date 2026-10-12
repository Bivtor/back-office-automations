param(
    [Parameter(Mandatory=$true)]
    [string]$CsvPath,
    [string]$Delimiter = ","
)

$errors = @()
$lineNumber = 0
$expectedCount = $null

try {
    $lines = Get-Content -Path $CsvPath -ErrorAction Stop
} catch {
    Write-Error "Could not read file: $_"
    exit 1
}

foreach ($line in $lines) {
    $lineNumber++

    if ([string]::IsNullOrWhiteSpace($line)) {
        $errors += "Line $lineNumber : Empty line"
        continue
    }

    # Count raw delimiters — no quote handling
    $count = ($line.ToCharArray() | Where-Object { $_ -eq $Delimiter }).Count

    if ($lineNumber -eq 1) {
        $expectedCount = $count
        Write-Host "Header row has $($count + 1) columns ($count delimiters)." -ForegroundColor Cyan
        continue
    }

    if ($count -ne $expectedCount) {
        $diff = $count - $expectedCount
        $direction = if ($diff -gt 0) { "TOO MANY (+$diff)" } else { "TOO FEW ($diff)" }
        $errors += "Line $lineNumber : $direction commas (got $count, expected $expectedCount) --> $line"
    }
}

Write-Host "`n--- CSV Shape Check Results ---" -ForegroundColor Yellow
Write-Host "File   : $CsvPath"
Write-Host "Lines  : $lineNumber"

if ($errors.Count -eq 0) {
    Write-Host "Status : No errors found!" -ForegroundColor Green
} else {
    Write-Host "Status : $($errors.Count) error(s) found" -ForegroundColor Red
    $errors | ForEach-Object { Write-Host "  ERROR: $_" -ForegroundColor Red }
}
