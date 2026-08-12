<#
.SYNOPSIS
    Turns a numbered caption .txt into one Audacity label track per category.

.DESCRIPTION
    Bound to a hotkey by "Audacity annotation hotkey.ahk". The annotation site
    hands out files like "Hut R0_sound_nouns.txt" where every caption carries the
    category it belongs to:

        0.000<TAB>30.000<TAB>8 - Insect chirping<TAB>Insect Chirping
        0.200<TAB>0.933<TAB>1 - Male speech<TAB>Male Speech

    Audacity's own File > Import > Labels flattens all of that into one track.
    This builds a track per number instead - named "1 Male speech", "2 Breathing"
    and so on from the optional fourth field, or just "1", "2", "3" when the file
    does not carry one - with the "N - " prefix stripped from each caption, since
    the track already says it.

    Audacity has no scriptable "import labels from this path" (ImportLabels only
    opens a dialog, and Import2 rejects label files), so the tracks are built
    label by label over mod-script-pipe.

.PARAMETER Path
    Skip the file picker and read this .txt.

.PARAMETER Replace
    Answer the existing-label-tracks prompt with "replace", unattended: the
    label tracks already there are deleted first.

.PARAMETER KeepExisting
    Answer it with "add alongside", unattended: the label tracks already there
    are left as they are and the new ones land beneath them.

.PARAMETER AnyProject
    Skip the check that the caption file belongs to the project Audacity has
    open. Without this, "Hut R0_sound_nouns.txt" refuses to import into a
    project called "Mud".

.PARAMETER DryRun
    Print the commands that would be sent and change nothing.
#>
[CmdletBinding()]
param(
    [string]$Path,
    [switch]$Replace,
    [switch]$KeepExisting,
    [switch]$AnyProject,
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot 'lib\AudacityPipe.ps1')
. (Join-Path $PSScriptRoot 'lib\SoundNouns.ps1')

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
# Show-Problem, Invoke-WithOwner and the focus fix behind them, shared with
# MakeSoundNouns.ps1. $DialogTitle below is what they put in the title bar.
. (Join-Path $PSScriptRoot 'lib\Dialogs.ps1')

$DialogTitle = 'Import sound nouns'

function ConvertTo-AudacityArg {
    <#
        Audacity's command parser reads Key="value", so a quote or backslash in
        the caption has to be escaped or the rest of the command is misread.
    #>
    param([string]$Value)
    $escaped = $Value -replace '\\', '\\'
    $escaped = $escaped -replace '"', '\"'
    return $escaped
}

function Format-AudacityTime {
    # Invariant, so a comma-decimal locale cannot turn 0.933 into "0,933".
    param([double]$Value)
    return [string]::Format([System.Globalization.CultureInfo]::InvariantCulture, '{0:0.000000}', $Value)
}

function Get-AudacityTrackList {
    $response = Invoke-AudacityCommands -Commands @('GetInfo: Type=Tracks Format=JSON')
    # Assign before wrapping. Windows PowerShell 5.1's ConvertFrom-Json pushes a
    # JSON array down the pipeline as ONE object, so @(... | ConvertFrom-Json)
    # yields a single element holding every track and .Count reads 1 - which
    # silently shifts every track index by one.
    $parsed = $response[0] | ConvertFrom-Json
    return , @($parsed)
}

function Get-CaptionFileProject {
    <#
        The project a caption file belongs to, taken from its name - whatever
        the download did to the spaces. The site's files arrive percent-encoded
        and sometimes lose the leading '%' of the first "%20", so the project
        name is simply whatever precedes the first separator:

            HUt%20R0_sound_nouns.txt   -> HUt
            Cast20-%20Annotations.txt  -> Cast
            Mud20-%20Completion.txt    -> Mud
            Pig R0_sound_nouns.txt     -> Pig
            Sub.txt                    -> Sub

        Splitting too eagerly is safe: this only decides whether to ask before
        importing, so the cost of guessing short is one extra confirmation.
    #>
    param([Parameter(Mandatory)][string]$Path)

    $raw = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    foreach ($pattern in '^(.+?)20-%20', '^(.+?)%20', '^(.+?)\s') {
        if ($raw -match $pattern) { return $Matches[1].Trim() }
    }
    return $raw.Trim()
}

function Get-AudacityLabelCount {
    <#
        Total labels in the project. Counted from the raw response rather than
        through Get-AudacityAnnotations because this is called right after the
        label tracks have been removed, when the reply can carry no entries at
        all and ConvertFrom-Json has nothing to chew on.
    #>
    $response = Invoke-AudacityCommands -Commands @('GetInfo: Type=Labels Format=JSON')
    $raw = $response[0]
    if (-not $raw -or $raw.Trim() -eq '') { return 0 }

    $parsed = $raw | ConvertFrom-Json
    if (-not $parsed) { return 0 }

    $count = 0
    foreach ($entry in $parsed) { $count += @($entry[1]).Count }
    return $count
}

try {
    $context = Get-AudacityProjectContext

    # --- which file -------------------------------------------------------
    if (-not $Path) {
        $dialog = New-Object System.Windows.Forms.OpenFileDialog
        $dialog.Title            = 'Import sound nouns from'
        $dialog.Filter           = 'Caption files (*.txt)|*.txt|All files (*.*)|*.*'
        $dialog.InitialDirectory = $context.Directory
        $dialog.CheckFileExists  = $true

        $guess = Find-SoundNounFile -Directory $context.Directory -ProjectName $context.Name
        if ($guess) { $dialog.FileName = [System.IO.Path]::GetFileName($guess) }

        $result = Invoke-WithOwner { param($owner) $dialog.ShowDialog($owner) }
        if ($result -ne [System.Windows.Forms.DialogResult]::OK) { exit 0 }
        $Path = $dialog.FileName
    }

    $groups = Read-SoundNounFile -Path $Path
    $labelTotal = ($groups | ForEach-Object { $_.Labels.Count } | Measure-Object -Sum).Sum
    $fileName = [System.IO.Path]::GetFileName($Path)

    # --- is this file even for this project? ------------------------------
    # Audacity's scripting pipe always talks to whichever project is frontmost,
    # so with several open it is easy to fire a file at the wrong one and
    # scribble over real annotation work. Check, and check again below.
    $wantProject = Get-CaptionFileProject -Path $Path
    if (-not $AnyProject -and $wantProject -and $context.Name -and
        $wantProject -ne $context.Name) {
        $message = "$fileName looks like it belongs to the project '$wantProject', " +
                   "but Audacity currently has '$($context.Name)' open.`n`n" +
                   "Import it into '$($context.Name)' anyway?"
        $answer = Invoke-WithOwner {
            param($owner)
            [System.Windows.Forms.MessageBox]::Show($owner, $message, $DialogTitle,
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Warning,
                [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
        }
        if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { exit 0 }
    }

    # --- existing label tracks --------------------------------------------
    $tracks = Get-AudacityTrackList
    $existing = @($tracks | Where-Object { $_.kind -eq 'label' })

    if ($existing.Count -gt 0) {
        $wipe = $false
        if ($Replace)          { $wipe = $true }
        elseif ($KeepExisting) { $wipe = $false }
        else {
            $names = ($existing | ForEach-Object { $_.name }) -join ', '
            $message = "This project already has $($existing.Count) label track(s):`n`n$names`n`n" +
                       # Braces round the name: '?' is a legal character in a
                       # PowerShell variable name, so "$fileName?" reads as a
                       # variable called 'fileName?' and the prompt loses both
                       # the filename and its question mark.
                       "What should the $($groups.Count) track(s) from ${fileName} do?`n`n" +
                       "Yes     - replace them: the existing label tracks are deleted first`n" +
                       "No      - add alongside: keep the existing label tracks as well`n" +
                       "Cancel  - leave this project alone"
            # Adding alongside is the default: Enter should not be able to delete
            # somebody's annotation work, and Escape still cancels outright.
            $answer = Invoke-WithOwner {
                param($owner)
                [System.Windows.Forms.MessageBox]::Show($owner, $message, $DialogTitle,
                    [System.Windows.Forms.MessageBoxButtons]::YesNoCancel,
                    [System.Windows.Forms.MessageBoxIcon]::Question,
                    [System.Windows.Forms.MessageBoxDefaultButton]::Button2)
            }
            switch ($answer) {
                ([System.Windows.Forms.DialogResult]::Yes) { $wipe = $true }
                ([System.Windows.Forms.DialogResult]::No)  { $wipe = $false }
                default                                    { exit 0 }
            }
        }

        if ($wipe) {
            # Select every label track, then remove in one go so it is a single
            # undo step. Track= is a 0-based index across tracks of all kinds.
            $commands = New-Object System.Collections.Generic.List[string]
            $mode = 'Set'
            for ($i = 0; $i -lt $tracks.Count; $i++) {
                if ($tracks[$i].kind -ne 'label') { continue }
                $commands.Add("SelectTracks: Track=$i TrackCount=1 Mode=$mode")
                $mode = 'Add'
            }
            $commands.Add('RemoveTracks:')

            if ($DryRun) {
                Write-Host "-- would remove $($existing.Count) label track(s):"
                $commands | ForEach-Object { Write-Host "   $_" }
                # Pretend they are gone so the indices below read as they would.
                $tracks = @($tracks | Where-Object { $_.kind -ne 'label' })
            }
            else {
                Invoke-AudacityCommands -Commands $commands | Out-Null
                $tracks = Get-AudacityTrackList
            }
        }
    }

    # --- build the tracks --------------------------------------------------
    # SetLabel's index is global across every label track in the project, so
    # note how many labels exist before adding any. New tracks land at the
    # bottom and labels go in ascending start order, so each label we add is the
    # newest one overall and its index is simply the running count.
    # Label tracks already present, so verification can ignore them and look
    # only at what this run created.
    $baseLabelTracks = @($tracks | Where-Object { $_.kind -eq 'label' }).Count
    # No label tracks means no labels - no need to ask, and it keeps the dry run
    # honest after it has pretended to remove them.
    if ($baseLabelTracks -eq 0) { $labelIndex = 0 } else { $labelIndex = Get-AudacityLabelCount }
    $trackIndex = $tracks.Count

    $commands = New-Object System.Collections.Generic.List[string]
    # What each group is about to become. Built here rather than recomputed
    # during verification, so the name sent to Audacity and the name reported
    # afterwards cannot drift apart.
    $plan     = New-Object System.Collections.Generic.List[object]
    $ordinal  = 0

    foreach ($group in $groups) {
        # "2 Breathing" where the file named the category, plain "2" where it
        # did not. The name comes from a workbook column, so it can hold spaces,
        # a trailing period, or a quote - it goes through the same escaping as
        # the caption text.
        $ordinal++
        $trackName = if ($group.Name) { "$($group.Number) $($group.Name)" } else { [string]$group.Number }
        # Label tracks are numbered 1..n among themselves and each new one lands
        # at the bottom, so this run owns the numbers just above the count we
        # started with, in group order. Their category numbers do not come into
        # it - a gap in the numbering shifts nothing.
        $plan.Add([pscustomobject]@{
            Group       = $group
            Name        = $trackName
            TrackNumber = $baseLabelTracks + $ordinal
        })

        $commands.Add('NewLabelTrack:')
        # Select AND focus it. AddLabel follows the *focused* track, not the
        # selected one - selecting alone leaves focus whereever it was and the
        # labels land in somebody else's track.
        $commands.Add("SelectTracks: Track=$trackIndex TrackCount=1 Mode=Set")
        $commands.Add('SetTrack: Focused=1')
        $commands.Add("SetTrack: Name=`"$(ConvertTo-AudacityArg $trackName)`"")

        foreach ($label in $group.Labels) {
            $start = Format-AudacityTime $label.Start
            $end   = Format-AudacityTime $label.End
            $text  = ConvertTo-AudacityArg $label.Text
            $commands.Add("SelectTime: Start=$start End=$end")
            $commands.Add('AddLabel:')
            $commands.Add("SetLabel: Label=$labelIndex Text=`"$text`"")
            $labelIndex++
        }
        $trackIndex++
    }
    $commands.Add('SelectNone:')

    if ($DryRun) {
        Write-Host "-- would send $($commands.Count) commands for $($groups.Count) tracks, $labelTotal labels:"
        $commands | ForEach-Object { Write-Host "   $_" }
        Write-Host "-- dry run, nothing was sent"
        exit 0
    }

    # Last look before anything is written. Everything above - the picker, the
    # replace prompt - gives the user time to switch projects in Audacity, and
    # the track indices computed above would then point into a different
    # project entirely.
    $now = Get-AudacityProjectContext
    if ($now.Name -ne $context.Name) {
        throw ("Audacity switched from '$($context.Name)' to '$($now.Name)' while this was " +
               "getting ready.`n`nNothing was imported. Bring '$($context.Name)' back to the " +
               "front and try again.")
    }

    Invoke-AudacityCommands -Commands $commands | Out-Null

    # --- verify -------------------------------------------------------------
    # Cheap, and it is the safety net for the index arithmetic above: a silent
    # off-by-one would file captions on the wrong track.
    # Assign before filtering. Get-AudacityAnnotations hands back all its rows
    # as ONE array object, so piping the call straight into Where-Object gives
    # it a single item to test instead of one per label - the same trap as
    # ConvertFrom-Json above.
    $allAfter = Get-AudacityAnnotations
    # TrackNumber is 1-based among label tracks, so anything above the count we
    # started with is ours - a kept track that happens to be called "3" cannot
    # be mistaken for the one just built.
    $after = @($allAfter | Where-Object { $_.TrackNumber -gt $baseLabelTracks })
    $problems = New-Object System.Collections.Generic.List[string]

    # Matched on TrackNumber, not on the name: two categories are allowed to
    # carry the same name, and matching on that would then compare a track
    # against both of them. The name still goes into every message, since that
    # is what the track is called on screen.
    foreach ($entry in $plan) {
        $group = $entry.Group
        $where = "track $($entry.TrackNumber) '$($entry.Name)'"
        $got = @($after | Where-Object { $_.TrackNumber -eq $entry.TrackNumber } | Sort-Object Start, End)
        if ($got.Count -ne $group.Labels.Count) {
            $problems.Add("  ${where}: expected $($group.Labels.Count) labels, found $($got.Count)")
            continue
        }
        for ($i = 0; $i -lt $got.Count; $i++) {
            $want = $group.Labels[$i]
            if ($got[$i].Text -ne $want.Text) {
                $problems.Add("  $where label $($i + 1): expected '$($want.Text)', found '$($got[$i].Text)'")
            }
            elseif ([math]::Abs($got[$i].Start - $want.Start) -gt 0.001 -or
                    [math]::Abs($got[$i].End   - $want.End)   -gt 0.001) {
                $problems.Add(("  {0} label {1}: expected {2:0.000}-{3:0.000}, found {4:0.000}-{5:0.000}" -f `
                    $where, ($i + 1), $want.Start, $want.End, $got[$i].Start, $got[$i].End))
            }
        }
    }

    if ($problems.Count -gt 0) {
        Show-Problem ("$fileName imported, but the result does not match the file:`n`n" +
                      (($problems | Select-Object -First 12) -join "`n") +
                      "`n`nUndo in Audacity (Ctrl+Z) and try again.")
        exit 1
    }

    Write-Host "$fileName  ->  $($groups.Count) tracks, $labelTotal labels"
}
catch {
    Show-Problem $_.Exception.Message
    exit 1
}
