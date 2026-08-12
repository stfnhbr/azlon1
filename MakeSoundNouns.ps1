<#
.SYNOPSIS
    Turns an annotation workbook into label tracks in Audacity, in one press.

.DESCRIPTION
    Bound to a hotkey by "Audacity annotation hotkey.ahk". Pick the .xlsx and
    this asks Claude for one sound-event noun per caption, builds the numbered
    caption file from those nouns, and hands it to ImportSoundNouns.ps1.

    The model is asked for the noun and nothing else. Start, end, the category
    number and the track name are copied from the workbook by this script, so
    they cannot drift from the source - which is what went wrong with the files
    made by hand, where fields three and four arrived swapped and the import
    refused the lot.

    Everything checkable is checked before the model is called, and the finished
    file goes through lib\SoundNouns.ps1 - the same parser the import uses -
    before anything is sent to Audacity.

.PARAMETER Path
    Skip the file picker and read this .xlsx.

.PARAMETER Sheet
    Use this worksheet instead of asking. Case sensitive: these names sometimes
    differ only by a trailing space.

.PARAMETER Model
    Which Claude to ask. Defaults to claude-opus-5.

.PARAMETER OutPath
    Write the caption file here instead of beside the workbook.

.PARAMETER NoImport
    Write the caption file and stop, leaving Audacity alone.

.PARAMETER Replace
    Passed through to the import: answer its "replace existing label tracks?"
    prompt with yes, unattended.

.PARAMETER DryRun
    Print the prompt and the path that would be written, and call nothing.
#>
[CmdletBinding()]
param(
    [string]$Path,
    [string]$Sheet,
    [string]$Model = 'claude-opus-5',
    [string]$OutPath,
    [switch]$NoImport,
    [switch]$Replace,
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib\AudacityPipe.ps1')
. (Join-Path $PSScriptRoot 'lib\SoundNouns.ps1')
. (Join-Path $PSScriptRoot 'lib\AnnotationWorkbookReader.ps1')
. (Join-Path $PSScriptRoot 'lib\ClaudeCli.ps1')

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
. (Join-Path $PSScriptRoot 'lib\Dialogs.ps1')

$DialogTitle = 'Sound nouns'

# How many captions go up in one call. Small enough that a reply stays reliable
# and a retry is cheap, large enough that a normal task is one or two calls.
$BatchSize = 150

# En-dash, as a char rather than a literal: Windows PowerShell 5.1 reads a
# BOM-less .ps1 as ANSI and would corrupt one written into the source.
$Dash = [char]0x2013

function Find-AnnotationWorkbook {
    <#
        Best guess at the workbook belonging to a project, or $null. The site's
        downloads are not consistently cased or spaced, so this only has to be
        close enough to pre-select something in the picker.
    #>
    param([string]$Directory, [string]$ProjectName)

    if (-not $Directory -or -not (Test-Path -LiteralPath $Directory)) { return $null }
    $candidates = @(Get-ChildItem -LiteralPath $Directory -Filter '*.xlsx' -File -ErrorAction SilentlyContinue)
    if ($candidates.Count -eq 0) { return $null }

    $match = $candidates |
        Where-Object { $_.Name.StartsWith($ProjectName, [System.StringComparison]::OrdinalIgnoreCase) } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    if ($match) { return $match.FullName }
    return $null
}

function Select-Sheet {
    <# A list of sheet names, for the workbooks that carry more than one. #>
    param([string[]]$Names)

    Invoke-WithOwner {
        param($owner)

        $form = New-Object System.Windows.Forms.Form
        $form.Text = $DialogTitle
        $form.FormBorderStyle = 'FixedDialog'
        $form.StartPosition = 'CenterScreen'
        $form.MinimizeBox = $false; $form.MaximizeBox = $false
        $form.ClientSize = New-Object System.Drawing.Size(320, 220)

        $label = New-Object System.Windows.Forms.Label
        $label.Text = 'Which sheet holds the captions?'
        $label.SetBounds(12, 12, 296, 18)
        $form.Controls.Add($label)

        $list = New-Object System.Windows.Forms.ListBox
        $list.SetBounds(12, 36, 296, 130)
        foreach ($n in $Names) { [void]$list.Items.Add($n) }
        $list.SelectedIndex = 0
        $form.Controls.Add($list)

        $ok = New-Object System.Windows.Forms.Button
        $ok.Text = 'OK'; $ok.SetBounds(146, 178, 75, 26)
        $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Controls.Add($ok); $form.AcceptButton = $ok

        $cancel = New-Object System.Windows.Forms.Button
        $cancel.Text = 'Cancel'; $cancel.SetBounds(233, 178, 75, 26)
        $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
        $form.Controls.Add($cancel); $form.CancelButton = $cancel

        # Double-click a name to take it.
        $list.Add_DoubleClick({ $form.DialogResult = [System.Windows.Forms.DialogResult]::OK }.GetNewClosure())

        try {
            if ($form.ShowDialog($owner) -ne [System.Windows.Forms.DialogResult]::OK) { return $null }
            return [string]$list.SelectedItem
        }
        finally { $form.Dispose() }
    }
}

function New-NounPrompt {
    <#
        One prompt: the project's noun rules, then the rows. The rules come
        straight out of Sound_Noun.md so there is one copy of them - the file is
        also what gets pasted into Claude on the web, and two copies would drift.
    #>
    param(
        [Parameter(Mandatory)]$Rows,
        [Parameter(Mandatory)][string]$Rules
    )

    $extras = @()
    if ($Rows | Where-Object { $_.SourceDescription }) { $extras += 'SourceDescription' }
    if ($Rows | Where-Object { $_.Prominence })        { $extras += 'Prominence' }

    $header = @('Caption', 'Track', 'TrackDescription', 'AnnotationText') + $extras

    $lines = New-Object System.Collections.Generic.List[string]
    $lines.Add(($header -join "`t"))
    foreach ($row in $Rows) {
        $fields = @(
            [string]$row.CaptionNumber
            ($row.Track            -replace '\s+', ' ')
            ($row.TrackDescription -replace '\s+', ' ')
            ($row.Text             -replace '\s+', ' ')
        )
        if ($extras -contains 'SourceDescription') { $fields += ($row.SourceDescription -replace '\s+', ' ') }
        if ($extras -contains 'Prominence')        { $fields += ($row.Prominence        -replace '\s+', ' ') }
        $lines.Add(($fields -join "`t"))
    }

    return @"
You are naming the sound event in each row of an audio-annotation sheet.

Reply with a JSON array and nothing else - no prose before or after it, no code
fence. One object per row, in the order the rows are given:

[{"caption": <the Caption value, unchanged>, "noun": "<sound-event noun>"}]

The noun is the only thing you choose. Start time, end time, track number and
track name are copied from the workbook by the script that called you, so never
put a number, a dash or a time inside "noun" - just the sound event, normally
one to three words.

Return one object for every row below, including rows whose noun repeats.

===== SOUND-NOUN RULES =====
$Rules
===== END OF RULES =====

The rules above also cover output format, file naming and how to reply in chat.
Those parts do not apply here: you are being called as a subroutine, the caller
builds the file, and this reply is the JSON array described at the top.

===== ROWS (tab separated) =====
$($lines -join "`n")
"@
}

function Get-NounMap {
    <#
    .SYNOPSIS
        Asks Claude for the nouns of one batch and returns captionNumber -> noun.
    .DESCRIPTION
        Everything the model sends back is checked against the rows that were
        asked about: one noun per caption, no caption invented or dropped, and
        nothing in the noun that belongs to the caller. A failing batch is asked
        once more with the specific complaints attached, since the second try
        usually lands, and only then does it give up.
    #>
    param(
        [Parameter(Mandatory)]$Rows,
        [Parameter(Mandatory)][string]$Rules,
        [string]$Model,
        [string]$Label = '',
        [scriptblock]$Tick
    )

    $prompt = New-NounPrompt -Rows $Rows -Rules $Rules
    $wanted = @{}
    foreach ($row in $Rows) { $wanted[[int]$row.CaptionNumber] = $true }

    for ($attempt = 1; $attempt -le 2; $attempt++) {
        $reply = Invoke-ClaudeJson -Prompt $prompt -Model $Model -Tick $Tick
        # Assign before wrapping: 5.1 pushes a JSON array down the pipeline as
        # one object, so @(... | ConvertFrom-Json) would collapse the lot.
        $parsed = ConvertFrom-ClaudeJsonBlock -Text $reply
        $items = @($parsed)

        $map = @{}
        $problems = New-Object System.Collections.Generic.List[string]

        foreach ($item in $items) {
            $caption = $null
            try { $caption = [int]$item.caption } catch { }
            if ($null -eq $caption -or -not $wanted.ContainsKey($caption)) {
                $problems.Add("caption '$($item.caption)' was not one of the rows")
                continue
            }
            if ($map.ContainsKey($caption)) {
                $problems.Add("caption $caption came back more than once")
                continue
            }

            $noun = ([string]$item.noun).Trim()
            if ($noun -eq '')                    { $problems.Add("caption ${caption}: the noun is empty") }
            elseif ($noun -match "[`t`r`n]")     { $problems.Add("caption ${caption}: the noun contains a tab or a line break") }
            elseif ($noun.Length -gt 60)         { $problems.Add("caption ${caption}: '$noun' is too long to be a sound noun") }
            elseif ($noun -match '^\s*[0-9]+\s*[\u2013\u2014-]') { $problems.Add("caption ${caption}: '$noun' still carries a number prefix") }
            else                                 { $map[$caption] = $noun }
        }

        foreach ($caption in ($wanted.Keys | Sort-Object)) {
            if (-not $map.ContainsKey($caption)) { $problems.Add("caption ${caption}: no noun came back") }
        }

        if ($problems.Count -eq 0) { return $map }

        if ($attempt -ge 2) {
            throw ("Claude's nouns$Label do not fit the rows, twice running:`n" +
                   (($problems | Select-Object -First 12) -join "`n"))
        }

        $prompt = $prompt + @"


===== YOUR PREVIOUS REPLY WAS REJECTED =====
$(($problems | Select-Object -First 20) -join "`n")

Send the whole array again, one object per row above, "caption" exactly as given
and "noun" a short sound event with no number, dash or time in it.
"@
    }
}

$script:progress = $null

try {
    $context = Get-AudacityProjectContext

    # --- which workbook ----------------------------------------------------
    if (-not $Path) {
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Title            = 'Make sound nouns from'
        $dialog.Filter           = 'Excel Workbook (*.xlsx)|*.xlsx|All files (*.*)|*.*'
        $dialog.InitialDirectory = $context.Directory
        $dialog.CheckFileExists  = $true

        $guess = Find-AnnotationWorkbook -Directory $context.Directory -ProjectName $context.Name
        if ($guess) { $dialog.FileName = [System.IO.Path]::GetFileName($guess) }

        $result = Invoke-WithOwner { param($owner) $dialog.ShowDialog($owner) }
        if ($result -ne [System.Windows.Forms.DialogResult]::OK) { exit 0 }
        $Path = $dialog.FileName
    }

    # --- read it -----------------------------------------------------------
    # The window goes up first: starting Excel and opening a workbook takes the
    # best part of ten seconds from cold, and the hotkey starts this process
    # hidden, so until something appears the keypress looks like it did nothing.
    $script:progress = New-ProgressWindow -Text "Reading $([System.IO.Path]::GetFileName($Path))..."
    $tick = { [System.Windows.Forms.Application]::DoEvents() }

    $book = Read-AnnotationWorkbook -Path $Path -Sheet $Sheet -ChooseSheet {
        param($names)
        # Out of the way of the picker, which is the one the user answers.
        if ($script:progress) { $script:progress.Close(); $script:progress = $null }
        Select-Sheet -Names $names
    }
    if (-not $book) { exit 0 }          # the sheet picker was cancelled
    if (-not $script:progress) { $script:progress = New-ProgressWindow -Text 'Reading the workbook...' }

    $rows = @($book.Rows)
    if ($rows.Count -eq 0) {
        throw "$([System.IO.Path]::GetFileName($book.Path)) sheet '$($book.Sheet)' holds no captions."
    }

    # --- is it fit to build a file from? -----------------------------------
    # Before the model call, not after: a workbook that cannot become label
    # tracks should say so straight away rather than after a minute of waiting.
    $problems = Test-AnnotationRows -Rows $rows
    if ($problems.Count -gt 0) {
        throw ("$([System.IO.Path]::GetFileName($book.Path)) sheet '$($book.Sheet)' cannot be turned " +
               "into label tracks:`n" + (($problems | Select-Object -First 12) -join "`n") +
               "`n`nFix the sheet and try again.")
    }

    if (-not $OutPath) {
        $stem = [System.IO.Path]::GetFileNameWithoutExtension($book.Path)
        $safe = ($book.Sheet.Trim() -replace '[\\/:\*\?"<>\|]', '-')
        $OutPath = Join-Path ([System.IO.Path]::GetDirectoryName($book.Path)) "$stem - $safe Sound Nouns.txt"
    }

    $rulesPath = Join-Path $PSScriptRoot 'Sound_Noun.md'
    if (-not (Test-Path -LiteralPath $rulesPath)) {
        throw "Cannot find the sound-noun rules:`n$rulesPath"
    }
    # Drop the skill frontmatter; it is metadata for the website, not guidance.
    $rules = [System.IO.File]::ReadAllText($rulesPath, [System.Text.Encoding]::UTF8)
    $rules = $rules -replace '(?s)\A---\r?\n.*?\r?\n---\r?\n', ''

    # --- batches -----------------------------------------------------------
    $batches = New-Object System.Collections.Generic.List[object]
    for ($i = 0; $i -lt $rows.Count; $i += $BatchSize) {
        $take = [Math]::Min($BatchSize, $rows.Count - $i)
        $batches.Add(@($rows[$i..($i + $take - 1)]))
    }

    if ($DryRun) {
        if ($script:progress) { $script:progress.Close(); $script:progress = $null }
        Write-Host "-- $($book.Sheet): $($rows.Count) captions in $($batches.Count) batch(es)"
        Write-Host "-- would write $OutPath"
        Write-Host "-- prompt for batch 1:"
        Write-Host (New-NounPrompt -Rows $batches[0] -Rules $rules)
        Write-Host "-- dry run, nothing was sent"
        exit 0
    }

    # --- ask ---------------------------------------------------------------
    $nouns = @{}
    for ($b = 0; $b -lt $batches.Count; $b++) {
        $label = if ($batches.Count -gt 1) { " (batch $($b + 1) of $($batches.Count))" } else { '' }
        $script:progress.SetText("Asking Claude for sound nouns$label...")
        $map = Get-NounMap -Rows $batches[$b] -Rules $rules -Model $Model -Label $label -Tick $tick
        foreach ($key in $map.Keys) { $nouns[$key] = $map[$key] }
    }
    $script:progress.SetText('Building the caption file...')

    # --- build it ----------------------------------------------------------
    # Only the noun came from Claude. Everything else is the workbook's own.
    $lines = New-Object System.Collections.Generic.List[string]
    foreach ($row in $rows) {
        $noun = $nouns[[int]$row.CaptionNumber]
        $lines.Add(("{0}`t{1}`t{2} {3} {4}`t{5}" -f
            [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, '{0:0.000}', $row.Start),
            [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, '{0:0.000}', $row.End),
            $row.TrackNumber, $Dash, $noun, $row.Track))
    }

    # UTF-8 with no BOM: Read-SoundNounFile reads the bytes as UTF-8 explicitly,
    # and a BOM would ride into the first line's start time.
    [System.IO.File]::WriteAllText($OutPath, ($lines -join "`r`n") + "`r`n",
        (New-Object System.Text.UTF8Encoding $false))

    # --- and check it the way the import will ------------------------------
    $groups = Read-SoundNounFile -Path $OutPath
    $labelTotal = ($groups | ForEach-Object { $_.Labels.Count } | Measure-Object -Sum).Sum
    Write-Host "$OutPath  ($($groups.Count) tracks, $labelTotal labels)"

    if ($NoImport) { exit 0 }

    # --- hand it to the import ---------------------------------------------
    # Out of the way first: the import has prompts of its own, and a topmost
    # window sitting over the one asking about existing label tracks would be a
    # poor place to leave the user.
    if ($script:progress) { $script:progress.Close(); $script:progress = $null }

    $importArgs = @{ Path = $OutPath }
    if ($Replace) { $importArgs['Replace'] = $true }
    & (Join-Path $PSScriptRoot 'ImportSoundNouns.ps1') @importArgs
    exit $LASTEXITCODE
}
catch {
    if ($script:progress) { $script:progress.Close(); $script:progress = $null }
    Show-Problem $_.Exception.Message
    exit 1
}
finally {
    if ($script:progress) { $script:progress.Close() }
}
