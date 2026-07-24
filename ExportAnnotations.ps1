<#
.SYNOPSIS
    Pulls the labels out of the running Audacity project and saves them as .xlsx.

.DESCRIPTION
    Bound to a hotkey by "Audacity annotation hotkey.ahk". Reads labels straight
    from Audacity over mod-script-pipe - no Export Labels step, and unlike the
    .txt export the source label track survives as a fourth column.

    Columns: Annotation | Start | End | Track, sorted by start time.

.PARAMETER OutPath
    Skip the save dialog and write here.

.PARAMETER NoOpen
    Write the workbook without opening it in Excel afterwards.
#>
[CmdletBinding()]
param(
    [string]$OutPath,
    [switch]$NoOpen
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib\AudacityPipe.ps1')
. (Join-Path $PSScriptRoot 'lib\AnnotationWorkbook.ps1')

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Show-Problem {
    param([string]$Message)
    [System.Windows.Forms.MessageBox]::Show($Message, 'Export annotations',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
}

try {
    $rows = Get-AudacityAnnotations
    if ($rows.Count -eq 0) {
        Show-Problem 'This Audacity project has no labels to export.'
        exit 1
    }

    $context = Get-AudacityProjectContext

    if (-not $OutPath) {
        $dialog = New-Object System.Windows.Forms.SaveFileDialog
        $dialog.Title            = 'Save annotations as'
        $dialog.Filter           = 'Excel Workbook (*.xlsx)|*.xlsx'
        $dialog.DefaultExt       = 'xlsx'
        $dialog.AddExtension     = $true
        $dialog.OverwritePrompt  = $true
        $dialog.InitialDirectory = $context.Directory
        $dialog.FileName         = "$($context.Name).xlsx"

        # PowerShell is launched hidden by the hotkey, so the dialog has no
        # natural owner and would open behind Audacity. A topmost stub form
        # gives it one and pulls it to the front.
        $owner = New-Object System.Windows.Forms.Form
        $owner.TopMost = $true; $owner.ShowInTaskbar = $false
        $owner.StartPosition = 'Manual'; $owner.Location = New-Object System.Drawing.Point(-2000, -2000)
        $owner.Size = New-Object System.Drawing.Size(1, 1)
        $owner.Show(); $owner.Activate()
        try {
            $result = $dialog.ShowDialog($owner)
        } finally {
            $owner.Close(); $owner.Dispose()
        }

        if ($result -ne [System.Windows.Forms.DialogResult]::OK) { exit 0 }
        $OutPath = $dialog.FileName
    }

    # The dialog already confirmed any overwrite, so honour the exact path asked
    # for rather than sliding to "name (2).xlsx".
    if (Test-Path -LiteralPath $OutPath) { Remove-Item -LiteralPath $OutPath -Force }

    $excel = $null
    $keepRunning = $false
    try {
        $excel = New-ExcelApp
        $written = Write-AnnotationWorkbook -Excel $excel -Rows $rows -Path $OutPath `
                                            -SheetName $context.Name
        if (-not $NoOpen) {
            Show-Workbooks -Excel $excel -Paths @($written)
            $keepRunning = $true
        }
    } finally {
        Close-ExcelApp -Excel $excel -KeepRunning:$keepRunning
    }

    Write-Host "$written  ($($rows.Count) labels)"
}
catch {
    Show-Problem $_.Exception.Message
    exit 1
}
