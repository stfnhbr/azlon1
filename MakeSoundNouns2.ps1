<#
.SYNOPSIS
    Turns an annotation workbook into label tracks in Audacity 3.7.9 or 4.0.

.DESCRIPTION
    Bound to Ctrl+Shift+Alt+N by "Audacity annotation hotkey.ahk". The second
    generation of MakeSoundNouns.ps1, which still works and is still the one
    used for 3.7.9 - the numbering follows CaptionFiller2/CaptionFiller3, and
    does not refer to an Audacity version.

    The expensive half of the job is the same on both versions and is not
    duplicated here: MakeSoundNouns.ps1 reads the workbook, asks Claude for one
    sound-event noun per caption, and writes the numbered caption file. What
    differs is only how those captions reach Audacity.

      3.7.9   the whole of MakeSoundNouns.ps1, unchanged, which hands the
              caption file to ImportSoundNouns.ps1 and builds the tracks live
              over mod-script-pipe.
      4.0     has no scripting interface of any kind, so the label tracks are
              written straight into the project file. See
              lib\AudacityProjectFile.ps1.

    That difference is worth knowing about before you press the key, because
    writing to the file means Audacity must not be holding it:

        save and close the project in Audacity 4.0 first.

    Everything else is the same. Tracks are named "2 Breathing" from the
    workbook's category number and name, appended at the bottom in order, in
    the same shape the pipe route produces - a project built either way reads
    back identically through Ctrl+Shift+Alt+K.

    The project file is copied to "NAME (before labels ...).aup4" before a byte
    is written, and the result is read back off the disk and counted before the
    run is called a success.

.PARAMETER Version
    Which Audacity to build for: 3, 4, or Auto. Auto follows the window in
    front, falls back to whichever version is running, and refuses to guess
    when both are open and neither is focused.

.PARAMETER ProjectPath
    Version 4 only. Write into this project file instead of asking. Accepts
    .aup4 and .aup3.

.PARAMETER CaptionPath
    Skip the workbook and the model entirely and import this caption file. This
    is also how to retry after a write failed without paying for the nouns
    twice - the caption file from the first run is already beside the workbook.

.PARAMETER ImportOnly
    Skip the workbook and the model and ask for a caption file to import. This
    is what Ctrl+Shift+Alt+I does, and it is here so that key keeps working on
    4.0 rather than quietly doing nothing there.

.PARAMETER Path
    Skip the file picker and read this .xlsx.

.PARAMETER Sheet
    Use this worksheet instead of asking. Case sensitive: these names sometimes
    differ only by a trailing space.

.PARAMETER Model
    Which Claude to ask. Defaults to claude-opus-5.

.PARAMETER NoCaptionNumbers
    Leave the workbook's caption number off the labels, so they read "Insect
    chirping" rather than "12 Insect chirping".

.PARAMETER Replace
    Answer the existing-label-tracks prompt with "replace", unattended.

.PARAMETER KeepExisting
    Answer it with "add alongside", unattended.

.PARAMETER AnyProject
    Do not ask when the caption file's name does not match the project.

.PARAMETER NoOpen
    Version 4 only. Do not offer to open the project when the write is done.

.PARAMETER DryRun
    Say what would be written, and write nothing.
#>
[CmdletBinding()]
param(
    [ValidateSet('3', '4', 'Auto')] [string]$Version = 'Auto',
    [string]$ProjectPath,
    [string]$CaptionPath,
    [switch]$ImportOnly,
    [string]$Path,
    [string]$Sheet,
    [string]$Model = 'claude-opus-5',
    [switch]$NoCaptionNumbers,
    [switch]$Replace,
    [switch]$KeepExisting,
    [switch]$AnyProject,
    [switch]$NoOpen,
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib\AudacityProjectFile.ps1')
. (Join-Path $PSScriptRoot 'lib\AudacityApps.ps1')
. (Join-Path $PSScriptRoot 'lib\SoundNouns.ps1')

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
. (Join-Path $PSScriptRoot 'lib\Dialogs.ps1')

$DialogTitle = 'Make sound nouns'

function Get-CaptionFileProject {
    <#
        The project a caption file belongs to, taken from its name - whatever
        the download did to the spaces. Same rule as ImportSoundNouns.ps1:

            HUt%20R0_sound_nouns.txt   -> HUt
            Cast20-%20Annotations.txt  -> Cast
            Pig R0_sound_nouns.txt     -> Pig

        Splitting too eagerly is safe: this only decides whether to ask before
        writing, so the cost of guessing short is one extra confirmation.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $raw = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    foreach ($pattern in '^(.+?)20-%20', '^(.+?)%20', '^(.+?)\s') {
        if ($raw -match $pattern) { return $Matches[1].Trim() }
    }
    return $raw.Trim()
}

function Show-Question {
    param([string]$Message, [string]$Buttons = 'YesNo', [string]$Icon = 'Question', [string]$Default = 'Button2')
    Invoke-WithOwner {
        param($owner)
        [System.Windows.Forms.MessageBox]::Show($owner, $Message, $DialogTitle,
            [System.Windows.Forms.MessageBoxButtons]::$Buttons,
            [System.Windows.Forms.MessageBoxIcon]::$Icon,
            [System.Windows.Forms.MessageBoxDefaultButton]::$Default)
    }
}

try {
    if ($ProjectPath) { $target = 4 } else { $target = Resolve-AudacityVersion -Version $Version }

    # ---- Audacity 3.7.9: the proven route, untouched ----------------------
    if ($target -eq 3) {
        $forward = @{}
        foreach ($name in 'Path', 'Sheet', 'Model') {
            if ($PSBoundParameters.ContainsKey($name)) { $forward[$name] = $PSBoundParameters[$name] }
        }
        foreach ($name in 'NoCaptionNumbers', 'Replace', 'KeepExisting', 'DryRun') {
            if ($PSBoundParameters.ContainsKey($name)) { $forward[$name] = $true }
        }
        if ($CaptionPath -or $ImportOnly) {
            # Nothing to make - the captions already exist, so this is an import.
            # Without a path, ImportSoundNouns.ps1 asks for one itself, which is
            # what Ctrl+Shift+Alt+I has always done.
            $importArgs = @{}
            if ($CaptionPath) { $importArgs['Path'] = $CaptionPath }
            foreach ($name in 'Replace', 'KeepExisting', 'AnyProject', 'DryRun') {
                if ($PSBoundParameters.ContainsKey($name)) { $importArgs[$name] = $true }
            }
            & (Join-Path $PSScriptRoot 'ImportSoundNouns.ps1') @importArgs
            exit $LASTEXITCODE
        }
        & (Join-Path $PSScriptRoot 'MakeSoundNouns.ps1') @forward
        exit $LASTEXITCODE
    }

    # ---- Audacity 4.0: build the file, then write into the project --------

    # The project is settled first, before the model is called: a run that
    # cannot be written anywhere should say so before it costs anything.
    if (-not $ProjectPath) {
        # Say so up front when 4.0 is holding something open, since the project
        # still open is usually the one that was meant. A warning rather than a
        # refusal: writing into a DIFFERENT project, one Audacity is not
        # holding, is perfectly fine and is why the picker still opens.
        $openNow = @(Get-Audacity4OpenProjects)
        if ($openNow.Count -gt 0) {
            $names = ($openNow | ForEach-Object { $_.Name }) -join ', '
            $answer = Show-Question (
                "Audacity 4.0 has $($openNow.Count) project open: $names.`n`n" +
                "Label tracks are written straight into the project file, which cannot be done " +
                "while Audacity is holding it - it would write its own copy back over them.`n`n" +
                "If one of those is the project you meant, save and close it in Audacity 4.0 " +
                "and press the hotkey again.`n`n" +
                "Carry on and pick a project that is not open?")
            if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { exit 0 }
        }

        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Title            = 'Write label tracks into which project?'
        $dialog.Filter           = 'Audacity projects (*.aup4;*.aup3)|*.aup4;*.aup3|All files (*.*)|*.*'
        $dialog.InitialDirectory = Get-Audacity4DefaultFolder
        $dialog.CheckFileExists  = $true
        $result = Invoke-WithOwner { param($owner) $dialog.ShowDialog($owner) }
        if ($result -ne [System.Windows.Forms.DialogResult]::OK) { exit 0 }
        $ProjectPath = $dialog.FileName
    }

    if (-not (Test-Path -LiteralPath $ProjectPath)) { throw "There is no project file at:`n$ProjectPath" }
    $ProjectPath = (Resolve-Path -LiteralPath $ProjectPath).ProviderPath
    if (Test-AudacityFileOpen -Path $ProjectPath) {
        throw ("Audacity still has this project open:`n$ProjectPath`n`n" +
               "Save and close it in Audacity 4.0, then press the hotkey again.")
    }

    $projectName = [System.IO.Path]::GetFileNameWithoutExtension($ProjectPath)
    $before = Read-AudacityProjectFile -Path $ProjectPath
    $existing = @($before.LabelTracks)

    # ---- the caption file -------------------------------------------------
    if ($CaptionPath) {
        if (-not (Test-Path -LiteralPath $CaptionPath)) { throw "There is no caption file at:`n$CaptionPath" }
        $captionFile = (Resolve-Path -LiteralPath $CaptionPath).ProviderPath
    }
    elseif ($ImportOnly) {
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Title            = "Import sound nouns into $projectName from"
        $dialog.Filter           = 'Caption files (*.txt)|*.txt|All files (*.*)|*.*'
        $dialog.InitialDirectory = [System.IO.Path]::GetDirectoryName($ProjectPath)
        $dialog.CheckFileExists  = $true

        $guess = Find-SoundNounFile -Directory $dialog.InitialDirectory -ProjectName $projectName
        if ($guess) { $dialog.FileName = [System.IO.Path]::GetFileName($guess) }

        $result = Invoke-WithOwner { param($owner) $dialog.ShowDialog($owner) }
        if ($result -ne [System.Windows.Forms.DialogResult]::OK) { exit 0 }
        $captionFile = $dialog.FileName
    }
    else {
        # MakeSoundNouns.ps1 does the whole of the expensive half and stops
        # before importing. It reports the file it wrote with Write-Host, which
        # in 5.1 is the information stream - hence 6>&1 to read it back.
        $forward = @{ NoImport = $true }
        foreach ($name in 'Path', 'Sheet', 'Model') {
            if ($PSBoundParameters.ContainsKey($name)) { $forward[$name] = $PSBoundParameters[$name] }
        }
        foreach ($name in 'NoCaptionNumbers', 'DryRun') {
            if ($PSBoundParameters.ContainsKey($name)) { $forward[$name] = $true }
        }

        $output = & (Join-Path $PSScriptRoot 'MakeSoundNouns.ps1') @forward 6>&1
        $made = $LASTEXITCODE
        if ($made -ne 0) { exit $made }      # it has already said what went wrong

        $captionFile = $null
        foreach ($line in @($output | ForEach-Object { [string]$_ })) {
            if ($line -match '^(?<p>.+\.txt)\s+\(\d+ tracks?, \d+ labels?\)\s*$') { $captionFile = $Matches['p'] }
        }
        if ($DryRun) {
            Write-Host "-- dry run: would write the caption file, then add its tracks to $ProjectPath"
            exit 0
        }
        # No path and a clean exit means a picker was cancelled, not a failure.
        if (-not $captionFile) { exit 0 }
        if (-not (Test-Path -LiteralPath $captionFile)) {
            throw ("The caption file was reported as`n$captionFile`nbut is not there. " +
                   "Nothing has been written to the project.")
        }
    }

    $groups = Read-SoundNounFile -Path $captionFile
    $labelTotal = ($groups | ForEach-Object { $_.Labels.Count } | Measure-Object -Sum).Sum
    $captionName = [System.IO.Path]::GetFileName($captionFile)

    # ---- is this file even for this project? ------------------------------
    $wantProject = Get-CaptionFileProject -Path $captionFile
    if (-not $AnyProject -and $wantProject -and $projectName -and $wantProject -ne $projectName) {
        $answer = Show-Question ("$captionName looks like it belongs to the project '$wantProject', " +
                                 "but you picked '$projectName'.`n`nWrite it into '$projectName' anyway?")
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { exit 0 }
    }

    # ---- existing label tracks --------------------------------------------
    $wipe = $false
    if ($existing.Count -gt 0) {
        if ($Replace)          { $wipe = $true }
        elseif ($KeepExisting) { $wipe = $false }
        else {
            $names = ($existing | ForEach-Object { if ($_.Name) { $_.Name } else { '(unnamed)' } }) -join ', '
            $message = "$projectName already has $($existing.Count) label track(s):`n`n$names`n`n" +
                       # Braces round the name: '?' is a legal character in a
                       # PowerShell variable name, so "$captionName?" reads as a
                       # variable and the prompt loses both name and question mark.
                       "What should the $($groups.Count) track(s) from ${captionName} do?`n`n" +
                       "Yes     - replace them: the existing label tracks are deleted first`n" +
                       "No      - add alongside: keep the existing label tracks as well`n" +
                       "Cancel  - leave this project alone"
            # Adding alongside is the default: Enter should not be able to delete
            # somebody's annotation work, and Escape still cancels outright.
            $answer = Show-Question $message -Buttons 'YesNoCancel'
            switch ($answer) {
                ([System.Windows.Forms.DialogResult]::Yes) { $wipe = $true }
                ([System.Windows.Forms.DialogResult]::No)  { $wipe = $false }
                default                                    { exit 0 }
            }
        }
    }

    if ($DryRun) {
        Write-Host "-- would write $($groups.Count) track(s), $labelTotal labels into $ProjectPath"
        Write-Host "-- existing label tracks: $($existing.Count), $(if ($wipe) { 'replaced' } else { 'kept' })"
        foreach ($group in $groups) {
            $trackName = if ($group.Name) { "$($group.Number) $($group.Name)" } else { [string]$group.Number }
            Write-Host ("   {0}  ({1} labels)" -f $trackName, $group.Labels.Count)
        }
        Write-Host '-- dry run, nothing was written'
        exit 0
    }

    # ---- write -------------------------------------------------------------
    $written = Add-AudacityProjectFileLabelTracks -Path $ProjectPath -Groups $groups -Replace:$wipe

    Write-Host "$captionName  ->  $ProjectPath  ($($written.TracksAdded) tracks, $labelTotal labels)"
    Write-Host "backup: $($written.BackupPath)"

    # Unlike the pipe route, nothing on screen changed - the project is closed.
    # So say so, rather than leaving a hotkey that appears to have done nothing.
    if (-not $NoOpen) {
        $summary = "$projectName now has $($written.LabelTracks) label track(s).`n`n" +
                   "Added $($written.TracksAdded) track(s), $labelTotal labels from $captionName."
        if ($written.TracksRemoved -gt 0) { $summary += "`nReplaced $($written.TracksRemoved) existing track(s)." }
        $summary += "`n`nThe project as it was is kept at:`n$($written.BackupPath)`n`nOpen it in Audacity 4.0 now?"

        $answer = Show-Question $summary -Icon 'Information' -Default 'Button1'
        if ($answer -eq [System.Windows.Forms.DialogResult]::Yes) {
            $exe = 'C:\Program Files\Audacity 4\bin\Audacity4.exe'
            if (Test-Path -LiteralPath $exe) { Start-Process -FilePath $exe -ArgumentList "`"$ProjectPath`"" }
            else { Show-Problem "Cannot find Audacity 4.0 at:`n$exe" }
        }
    }
}
catch {
    Show-Problem $_.Exception.Message
    exit 1
}
