<#
.SYNOPSIS
    Fixes the Excel "PERSONAL.XLSB is locked for editing by <you>" prompt.

.DESCRIPTION
    Runs the checks and fixes from the troubleshooting list, in order:
      1. Ends every hidden or leftover EXCEL.EXE process.
      2. Deletes stale ~$PERSONAL.XLSB lock files in the XLSTART folder.
      3. Turns off "Ignore other applications that use DDE" in Excel options.
      4. Reports duplicate PERSONAL.XLSB copies in other startup folders.
      5. Warns if the XLSTART folder is inside a OneDrive-synced location.
      6. Optionally marks PERSONAL.XLSB read-only (-MakeReadOnly) so Excel
         never asks again. Use -ClearReadOnly to undo that when you want to
         save new macros.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Fix-PersonalXlsb.ps1

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Fix-PersonalXlsb.ps1 -MakeReadOnly

.NOTES
    Save any open Excel work before running. Step 1 closes Excel without asking.
#>
[CmdletBinding()]
param(
    [switch]$MakeReadOnly,
    [switch]$ClearReadOnly,
    [switch]$SkipDde
)

$ErrorActionPreference = 'Continue'

function Write-Step($text) { Write-Host ""; Write-Host "== $text" -ForegroundColor Cyan }
function Write-Ok($text)   { Write-Host "   OK   $text" -ForegroundColor Green }
function Write-Warn2($text){ Write-Host "   WARN $text" -ForegroundColor Yellow }
function Write-Info($text) { Write-Host "   ..   $text" }

$xlstart  = Join-Path $env:APPDATA 'Microsoft\Excel\XLSTART'
$personal = Join-Path $xlstart 'PERSONAL.XLSB'

Write-Host "PERSONAL.XLSB lock fixer" -ForegroundColor White
Write-Host "User XLSTART folder: $xlstart"

# ---------------------------------------------------------------------------
Write-Step "1. Closing every Excel process"
$procs = Get-Process -Name EXCEL -ErrorAction SilentlyContinue
if ($procs) {
    Write-Info ("Found {0} EXCEL.EXE process(es). Ending them." -f $procs.Count)
    if ($procs.Count -gt 1) {
        Write-Warn2 "More than one Excel was running. That is the direct cause of the lock prompt."
    }
    $procs | Stop-Process -Force -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    Write-Ok "Excel closed."
} else {
    Write-Ok "No Excel process was running."
}

# ---------------------------------------------------------------------------
Write-Step "2. Removing stale lock files"
if (Test-Path $xlstart) {
    $locks = Get-ChildItem -Path $xlstart -Filter '~$*' -Force -File -ErrorAction SilentlyContinue
    if ($locks) {
        foreach ($l in $locks) {
            Remove-Item -LiteralPath $l.FullName -Force -ErrorAction SilentlyContinue
            Write-Ok "Deleted $($l.Name)"
        }
    } else {
        Write-Ok "No lock files found."
    }
} else {
    Write-Warn2 "XLSTART folder does not exist. Excel will create it when you record a macro to PERSONAL.XLSB."
}

if (Test-Path $personal) {
    Write-Ok "PERSONAL.XLSB found ($([math]::Round((Get-Item $personal).Length / 1KB)) KB)."
} else {
    Write-Warn2 "PERSONAL.XLSB is not in the user XLSTART folder. See step 4 for other copies."
}

# ---------------------------------------------------------------------------
Write-Step "3. Turning off 'Ignore other applications that use DDE'"
if ($SkipDde) {
    Write-Info "Skipped (-SkipDde)."
} else {
    $xl = $null
    try {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        $before = $xl.IgnoreRemoteRequests
        $xl.IgnoreRemoteRequests = $false
        $alt = $xl.AltStartupPath
        Write-Ok ("Setting was {0}, now False." -f $(if ($before) { 'ON (bad)' } else { 'already off' }))
        if ($alt) {
            Write-Warn2 "Excel also opens files from an alternate startup folder: $alt"
            $altCopy = Join-Path $alt 'PERSONAL.XLSB'
            if (Test-Path $altCopy) {
                Write-Warn2 "A second PERSONAL.XLSB lives there. Remove one copy or Excel opens it twice."
            }
        }
    } catch {
        Write-Warn2 "Could not talk to Excel through COM: $($_.Exception.Message)"
        Write-Info "Do it by hand: File > Options > Advanced > General > untick 'Ignore other applications that use DDE'."
    } finally {
        if ($xl) {
            try { $xl.Quit() } catch {}
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
            $xl = $null
            [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        }
        Get-Process -Name EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
    }
}

# ---------------------------------------------------------------------------
Write-Step "4. Looking for duplicate PERSONAL.XLSB copies"
$otherStartups = @(
    "$env:ProgramFiles\Microsoft Office\root\Office16\XLSTART",
    "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\XLSTART",
    "$env:ProgramFiles\Microsoft Office\Office16\XLSTART",
    "${env:ProgramFiles(x86)}\Microsoft Office\Office16\XLSTART",
    "$env:ProgramFiles\Microsoft Office\root\Office15\XLSTART",
    "${env:ProgramFiles(x86)}\Microsoft Office\root\Office15\XLSTART"
) | Where-Object { $_ -and (Test-Path $_) }

$dupes = @()
foreach ($dir in $otherStartups) {
    $c = Join-Path $dir 'PERSONAL.XLSB'
    if (Test-Path $c) { $dupes += $c }
}
if ($dupes) {
    foreach ($d in $dupes) { Write-Warn2 "Duplicate copy: $d" }
    Write-Info "Keep only the one in $xlstart. Move the others somewhere else (do not delete until you checked the macros inside)."
} else {
    Write-Ok "No duplicates in the Office program startup folders."
}

# ---------------------------------------------------------------------------
Write-Step "5. Checking cloud sync"
$syncRoots = @($env:OneDrive, $env:OneDriveConsumer, $env:OneDriveCommercial) | Where-Object { $_ }
$synced = $false
foreach ($root in $syncRoots) {
    if ($xlstart.StartsWith($root, [System.StringComparison]::OrdinalIgnoreCase)) { $synced = $true }
}
if ($synced) {
    Write-Warn2 "XLSTART is inside a OneDrive folder. Exclude it from sync, or the sync client can lock the file at startup."
} else {
    Write-Ok "XLSTART is not inside OneDrive."
}

# ---------------------------------------------------------------------------
Write-Step "6. Read-only attribute on PERSONAL.XLSB"
if (Test-Path $personal) {
    $item = Get-Item $personal
    if ($MakeReadOnly) {
        $item.IsReadOnly = $true
        Write-Ok "Set read-only. Excel opens it silently now. Run with -ClearReadOnly before saving new macros."
    } elseif ($ClearReadOnly) {
        $item.IsReadOnly = $false
        Write-Ok "Cleared read-only. You can save macros into PERSONAL.XLSB again."
    } else {
        Write-Info ("Currently read-only: {0}. Add -MakeReadOnly if the prompt keeps coming back." -f $item.IsReadOnly)
    }
}

Write-Host ""
Write-Host "Done. Start Excel normally and check whether the prompt is gone." -ForegroundColor White
Write-Host "If it comes back, open Task Manager while the prompt is on screen and count the EXCEL.EXE entries."
