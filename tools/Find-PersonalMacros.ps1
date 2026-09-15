<#
.SYNOPSIS
    Finds PERSONAL.XLSB and brings the personal macros back after they "disappeared".

.DESCRIPTION
    1. Shows whether PERSONAL.XLSB exists in the user XLSTART folder, its size and date.
    2. Searches the whole user profile for any PERSONAL*.xlsb copies, backups or autosaves.
    3. Lists Excel's "Disabled Items" (Excel disables a startup file after a crash or a
       forced close). With -ReEnable it clears that list so the file loads again.
    4. Opens the file through Excel and reports whether it still contains VBA code,
       plus the module names when Excel allows access to the VBA project.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Find-PersonalMacros.ps1

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Find-PersonalMacros.ps1 -ReEnable

.NOTES
    Close Excel before running. Nothing is deleted or overwritten.
#>
[CmdletBinding()]
param(
    [switch]$ReEnable,
    [switch]$SkipSearch
)

$ErrorActionPreference = 'Continue'
function Write-Step($t) { Write-Host ""; Write-Host "== $t" -ForegroundColor Cyan }
function Write-Ok($t)   { Write-Host "   OK   $t" -ForegroundColor Green }
function Write-Warn2($t){ Write-Host "   WARN $t" -ForegroundColor Yellow }
function Write-Info($t) { Write-Host "   ..   $t" }

$xlstart  = Join-Path $env:APPDATA 'Microsoft\Excel\XLSTART'
$personal = Join-Path $xlstart 'PERSONAL.XLSB'

# ---------------------------------------------------------------------------
Write-Step "1. The expected file"
if (Test-Path $personal) {
    $f = Get-Item $personal -Force
    Write-Ok "$personal"
    Write-Info ("Size: {0} KB   Last changed: {1}   Read-only: {2}" -f [math]::Round($f.Length / 1KB), $f.LastWriteTime, $f.IsReadOnly)
    if ($f.Length -lt 8KB) { Write-Warn2 "Very small. A PERSONAL.XLSB with macros is usually 10 KB or more. This may be a fresh empty one." }
} else {
    Write-Warn2 "Not found at $personal"
}

# ---------------------------------------------------------------------------
Write-Step "2. Every PERSONAL*.xlsb on this user profile (can take a minute)"
if ($SkipSearch) {
    Write-Info "Skipped (-SkipSearch)."
} else {
    $found = Get-ChildItem -Path $env:USERPROFILE -Recurse -Force -File -Include 'PERSONAL*.xlsb','PERSONAL*.xlsm','PERSONAL*.xls' -ErrorAction SilentlyContinue |
             Sort-Object LastWriteTime -Descending
    if ($found) {
        foreach ($x in $found) {
            Write-Info ("{0,8} KB  {1}  {2}" -f [math]::Round($x.Length / 1KB), $x.LastWriteTime.ToString('yyyy-MM-dd HH:mm'), $x.FullName)
        }
        Write-Info "The biggest or most recently changed copy is usually the one with your macros."
    } else {
        Write-Warn2 "No PERSONAL file found anywhere under $env:USERPROFILE."
    }
    $unsaved = Join-Path $env:LOCALAPPDATA 'Microsoft\Office\UnsavedFiles'
    if (Test-Path $unsaved) {
        $u = Get-ChildItem $unsaved -Force -File -ErrorAction SilentlyContinue
        if ($u) { Write-Info "Autorecovered files also exist in $unsaved" }
    }
}

# ---------------------------------------------------------------------------
Write-Step "3. Excel 'Disabled Items' list"
$anyDisabled = $false
foreach ($ver in '16.0','15.0','14.0') {
    $key = "HKCU:\Software\Microsoft\Office\$ver\Excel\Resiliency\DisabledItems"
    if (Test-Path $key) {
        $vals = Get-Item $key | Select-Object -ExpandProperty Property
        foreach ($v in $vals) {
            $anyDisabled = $true
            $raw = (Get-ItemProperty $key -Name $v).$v
            $text = ''
            try { $text = [System.Text.Encoding]::Unicode.GetString($raw) -replace '[^ -~ -ÿ]', ' ' } catch {}
            Write-Warn2 "Disabled item in Office $ver : $text"
            if ($ReEnable) {
                Remove-ItemProperty -Path $key -Name $v -ErrorAction SilentlyContinue
                Write-Ok "Re-enabled."
            }
        }
    }
}
if (-not $anyDisabled) {
    Write-Ok "Nothing is disabled."
} elseif (-not $ReEnable) {
    Write-Info "Run again with -ReEnable to clear this list, or in Excel: File > Options > Add-ins > Manage 'Disabled Items' > Go > Enable."
}

# ---------------------------------------------------------------------------
Write-Step "4. Does the file still contain macros?"
if (Test-Path $personal) {
    $xl = $null
    try {
        $xl = New-Object -ComObject Excel.Application
        $xl.Visible = $false
        $xl.DisplayAlerts = $false
        $wb = $null
        foreach ($w in $xl.Workbooks) { if ($w.Name -ieq 'PERSONAL.XLSB') { $wb = $w } }
        if ($wb) {
            Write-Ok "Excel loaded PERSONAL.XLSB at startup by itself, so it is not disabled right now."
        } else {
            Write-Warn2 "Excel did NOT load PERSONAL.XLSB at startup. Opening it directly to inspect."
            $wb = $xl.Workbooks.Open($personal, 0, $true)
        }
        if ($wb.HasVBProject) {
            Write-Ok "The file contains VBA code. Your macros are still inside."
            try {
                $names = @()
                foreach ($c in $wb.VBProject.VBComponents) { $names += ("{0} ({1} lines)" -f $c.Name, $c.CodeModule.CountOfLines) }
                foreach ($n in $names) { Write-Info $n }
            } catch {
                Write-Info "Module names hidden. To list them, enable File > Options > Trust Center > Trust Center Settings > Macro Settings > 'Trust access to the VBA project object model'."
            }
        } else {
            Write-Warn2 "The file has NO VBA code. It is an empty PERSONAL.XLSB. Look at the other copies in step 2 or restore from a backup."
        }
    } catch {
        Write-Warn2 "Could not inspect through Excel: $($_.Exception.Message)"
    } finally {
        if ($xl) {
            try { $xl.Quit() } catch {}
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
            [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        }
    }
}

Write-Host ""
Write-Host "Next: start Excel, press Alt+F8, set 'Macros in' to 'All Open Workbooks'." -ForegroundColor White
Write-Host "If PERSONAL.XLSB macros show there, you are done. If not, send me this output."
