<#
.SYNOPSIS
    Pulls the labels out of the running Audacity project and saves them as .xlsx.

.DESCRIPTION
    Bound to a hotkey by "Audacity annotation hotkey.ahk". Reads labels straight
    from Audacity over mod-script-pipe - no Export Labels step, and unlike
    Audacity's own .txt export the source label track survives as a column of
    its own.
    (Keep help lines from starting with ".txt": PowerShell reads a leading dot
    as a help keyword and silently drops the whole comment-based help block.)

    Ten columns in the "Soup EE" order - Caption Number, TrackNumber, Track,
    TrackDescription, StartTime(s), EndTime(s), SourceVisibility,
    SourceDescription, Prominence, AnnotationText - sorted by start time. The
    layout, including how to switch back to the old headers, lives in
    lib\AnnotationWorkbook.ps1.

.PARAMETER OutPath
    Skip the save dialog and write here.

.PARAMETER NoOpen
    Write the workbook without opening it in Excel afterwards.

.PARAMETER Layout
    Column layout for this one export: SoupEE or Legacy. Omit to follow the
    default set in lib\AnnotationWorkbook.ps1.
#>
[CmdletBinding()]
param(
    [string]$OutPath,
    [switch]$NoOpen,
    [ValidateSet('SoupEE', 'Legacy')] [string]$Layout
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib\AudacityPipe.ps1')
. (Join-Path $PSScriptRoot 'lib\AnnotationWorkbook.ps1')

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# Show-Problem, Invoke-WithOwner and the focus fix behind them, shared with the
# other two hotkey scripts. $DialogTitle below is what they put in the title bar.
. (Join-Path $PSScriptRoot 'lib\Dialogs.ps1')

$DialogTitle = 'Export annotations'

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
        # natural owner and would open behind Audacity. Invoke-WithOwner gives it
        # one and makes the activation stick - a stub form on its own is not
        # enough, since Windows refuses the foreground to a process that does not
        # already hold it, and the dialog then sits behind Audacity unseen.
        $result = Invoke-WithOwner { param($owner) $dialog.ShowDialog($owner) }
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
        # Omitted rather than defaulted, so the layout chosen in
        # lib\AnnotationWorkbook.ps1 stays the single source of truth.
        $layoutArg = @{}
        if ($Layout) { $layoutArg['Layout'] = $Layout }
        $written = Write-AnnotationWorkbook -Excel $excel -Rows $rows -Path $OutPath `
                                            -SheetName $context.Name @layoutArg
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
