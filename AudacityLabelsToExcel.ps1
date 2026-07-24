<#
.SYNOPSIS
    Converts Audacity label track exports (.txt) into Excel workbooks (.xlsx).

.DESCRIPTION
    Audacity exports labels as tab-separated lines:  start <TAB> end <TAB> label
    Optional frequency-range lines (starting with "\") are ignored.

    Each input file produces a sibling .xlsx with three columns:
        A: Annotation | B: Start | C: End

    For a live project, ExportAnnotations.ps1 is better - it reads Audacity
    directly and keeps the label track name too. Use this one for .txt files you
    already exported, or that came from someone else.

.PARAMETER Paths
    One or more label .txt files. Usually supplied by dropping files onto
    "Drop labels here.cmd".

.PARAMETER NoOpen
    Write the workbook(s) without opening them in Excel afterwards.
#>
[CmdletBinding()]
param(
    [Parameter(Position = 0, ValueFromRemainingArguments = $true)]
    [string[]]$Paths,

    [switch]$NoOpen
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib\AnnotationWorkbook.ps1')

$inv = [System.Globalization.CultureInfo]::InvariantCulture

function Read-AudacityLabels {
    param([string]$Path)

    $rows = New-Object System.Collections.Generic.List[object]
    $lineNo = 0

    foreach ($line in [System.IO.File]::ReadAllLines($Path, [System.Text.Encoding]::UTF8)) {
        $lineNo++
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        # Frequency-range continuation line belonging to the previous label.
        if ($line.StartsWith('\')) { continue }

        # Only the first two tabs are separators; the label itself may contain tabs.
        $parts = $line -split "`t", 3
        if ($parts.Count -lt 2) {
            Write-Warning "$([System.IO.Path]::GetFileName($Path)) line ${lineNo}: not tab-separated, skipped."
            continue
        }

        [double]$start = 0; [double]$end = 0
        $okStart = [double]::TryParse($parts[0].Trim(), 'Float', $inv, [ref]$start)
        $okEnd   = [double]::TryParse($parts[1].Trim(), 'Float', $inv, [ref]$end)
        if (-not ($okStart -and $okEnd)) {
            Write-Warning "$([System.IO.Path]::GetFileName($Path)) line ${lineNo}: unparseable times, skipped."
            continue
        }

        $text = if ($parts.Count -ge 3) { $parts[2].TrimEnd("`r") } else { '' }

        $rows.Add([pscustomobject]@{ Text = $text; Start = $start; End = $end })
    }

    return , $rows
}

# ---------------------------------------------------------------- validation --

if (-not $Paths -or $Paths.Count -eq 0) {
    Write-Host "Drag one or more Audacity label .txt files onto 'Drop labels here.cmd'." -ForegroundColor Yellow
    exit 1
}

$files = @()
foreach ($p in $Paths) {
    if (Test-Path -LiteralPath $p -PathType Leaf) {
        $files += (Resolve-Path -LiteralPath $p).ProviderPath
    } else {
        Write-Warning "Not a file, skipped: $p"
    }
}
if ($files.Count -eq 0) { Write-Host 'Nothing to convert.' -ForegroundColor Yellow; exit 1 }

# -------------------------------------------------------------------- convert --

$excel = $null
$written = @()
$keepRunning = $false
try {
    $excel = New-ExcelApp

    foreach ($file in $files) {
      try {
        Write-Host "Reading $([System.IO.Path]::GetFileName($file))..."
        $rows = Read-AudacityLabels -Path $file
        if ($rows.Count -eq 0) {
            Write-Warning "No labels found in $file - skipped."
            continue
        }

        $outPath = Write-AnnotationWorkbook -Excel $excel -Rows $rows `
                       -Path ([System.IO.Path]::ChangeExtension($file, '.xlsx'))

        $written += $outPath
        Write-Host "  -> $([System.IO.Path]::GetFileName($outPath))  ($($rows.Count) labels)" -ForegroundColor Green
      }
      catch {
        # One bad file must not take the rest of the batch down with it.
        Write-Warning "Failed to convert $([System.IO.Path]::GetFileName($file)): $($_.Exception.Message)"
      }
    }

    # Open them all in the instance that just wrote them, rather than shelling
    # out per file - see Show-Workbooks for why.
    if ($written.Count -gt 0 -and -not $NoOpen) {
        Show-Workbooks -Excel $excel -Paths $written
        $keepRunning = $true
    }
}
finally {
    Close-ExcelApp -Excel $excel -KeepRunning:$keepRunning
}

if ($written.Count -eq 0) { Write-Host 'No workbooks written.' -ForegroundColor Yellow; exit 1 }
