<#
.SYNOPSIS
    Shows every running Excel, when it started, and which program launched it.

.DESCRIPTION
    Run this WHILE the "PERSONAL.XLSB is locked" prompt is on screen. Do not click
    the prompt first. The output tells you whether one Excel or several are running,
    and for each one the parent program (explorer.exe = you double-clicked,
    python.exe = a script, OUTLOOK.EXE = an attachment, and so on) plus the
    command line it was started with.

.EXAMPLE
    powershell -ExecutionPolicy Bypass -File .\Who-StartedExcel.ps1
#>
$procs = Get-CimInstance Win32_Process -Filter "Name = 'EXCEL.EXE'"
if (-not $procs) { Write-Host "No Excel process is running."; return }

Write-Host ("{0} Excel process(es) running" -f @($procs).Count) -ForegroundColor Cyan
Write-Host ""
foreach ($p in $procs) {
    $parent = Get-CimInstance Win32_Process -Filter "ProcessId = $($p.ParentProcessId)" -ErrorAction SilentlyContinue
    $parentName = if ($parent) { $parent.Name } else { "(already exited, PID $($p.ParentProcessId))" }
    $gp = $null
    if ($parent) { $gp = Get-CimInstance Win32_Process -Filter "ProcessId = $($parent.ParentProcessId)" -ErrorAction SilentlyContinue }
    $win = (Get-Process -Id $p.ProcessId -ErrorAction SilentlyContinue).MainWindowTitle
    Write-Host ("PID {0}" -f $p.ProcessId) -ForegroundColor White
    Write-Host ("   started     : {0}" -f $p.CreationDate)
    Write-Host ("   window      : {0}" -f $(if ($win) { $win } else { "(none, hidden)" }))
    Write-Host ("   started by  : {0}" -f $parentName)
    if ($gp) { Write-Host ("   which was started by: {0}" -f $gp.Name) }
    Write-Host ("   command line: {0}" -f $p.CommandLine)
    Write-Host ("   path        : {0}" -f $p.ExecutablePath)
    Write-Host ""
}
Write-Host "A process with no window that started long before the others is the leftover one." -ForegroundColor Yellow
Write-Host "Two different 'path' values mean two Excel installations are fighting over the file." -ForegroundColor Yellow
