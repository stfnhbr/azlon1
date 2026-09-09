<#
.SYNOPSIS
    Pulls the labels out of Audacity 3.7.9 or Audacity 4.0 and saves them as .xlsx.

.DESCRIPTION
    Bound to Ctrl+Shift+Alt+K by "Audacity annotation hotkey.ahk". The second
    generation of ExportAnnotations.ps1, which still works and is still the one
    used for 3.7.9 - the numbering follows CaptionFiller2/CaptionFiller3, and
    does not refer to an Audacity version.

    Audacity 4.0 ships no scripting interface at all: no mod-script-pipe, no
    modules folder, nothing to talk to. So the two versions are reached two
    different ways, and which one is used is the only thing that changes here:

      3.7.9   over mod-script-pipe, exactly as before - the live project,
              including edits made since the last save.
      4.0     by reading the project file. 4.0 keeps unsaved work in the same
              database, in its autosave table, so this is the live project too
              rather than the last saved state. See lib\AudacityProjectFile.ps1.

    Either way the workbook is identical: ten columns in the "Soup EE" order,
    sorted by start time, with the label's own track as a column of its own.

.PARAMETER Version
    Which Audacity to read: 3, 4, or Auto. Auto follows the window in front,
    falls back to whichever version is running, and refuses to guess when both
    are open and neither is focused.

.PARAMETER ProjectPath
    Version 4 only. Read this project file instead of working out which one
    Audacity has open. Also reads a closed .aup4 or .aup3.

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
    [ValidateSet('3', '4', 'Auto')] [string]$Version = 'Auto',
    [string]$ProjectPath,
    [string]$OutPath,
    [switch]$NoOpen,
    [ValidateSet('SoupEE', 'Legacy')] [string]$Layout
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib\AudacityPipe.ps1')
. (Join-Path $PSScriptRoot 'lib\AudacityProjectFile.ps1')
. (Join-Path $PSScriptRoot 'lib\AudacityApps.ps1')
. (Join-Path $PSScriptRoot 'lib\AnnotationWorkbook.ps1')

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# Show-Problem, Invoke-WithOwner and the focus fix behind them, shared with the
# other hotkey scripts. $DialogTitle below is what they put in the title bar.
. (Join-Path $PSScriptRoot 'lib\Dialogs.ps1')

$DialogTitle = 'Export annotations'

try {
    # A path given outright settles the question; there is no point asking
    # which Audacity is in front when the file to read is already named.
    if ($ProjectPath) { $target = 4 } else { $target = Resolve-AudacityVersion -Version $Version }

    if ($target -eq 3) {
        $rows    = Get-AudacityAnnotations
        $context = Get-AudacityProjectContext
        $source  = 'Audacity 3.7.9'
    }
    else {
        if ($ProjectPath) {
            if (-not (Test-Path -LiteralPath $ProjectPath)) {
                throw "There is no project file at:`n$ProjectPath"
            }
            $project = [pscustomobject]@{
                Path  = (Resolve-Path -LiteralPath $ProjectPath).ProviderPath
                Saved = $true
                Name  = [System.IO.Path]::GetFileNameWithoutExtension($ProjectPath)
            }
        }
        else {
            $project = Select-Audacity4ProjectFile -Title $DialogTitle
            if (-not $project) { exit 0 }        # the picker was cancelled
        }

        $rows = Get-AudacityProjectFileAnnotations -Path $project.Path

        # A saved project names its own folder. One that has never been saved
        # lives in Audacity's SessionData, which is no place to offer to put a
        # workbook, so the folder 4.0 last used is offered instead.
        $folder = if ($project.Saved) { [System.IO.Path]::GetDirectoryName($project.Path) }
                  else                { Get-Audacity4DefaultFolder }
        $context = [pscustomobject]@{ Name = $project.Name; Directory = $folder }
        $source  = 'Audacity 4.0'
    }

    if ($rows.Count -eq 0) {
        Show-Problem "This $source project has no labels to export."
        exit 1
    }

    if (-not $OutPath) {
        $dialog = New-Object System.Windows.Forms.SaveFileDialog
        $dialog.Title            = "Save annotations as ($source)"
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

    Write-Host "$written  ($($rows.Count) labels from $source)"
}
catch {
    Show-Problem $_.Exception.Message
    exit 1
}
