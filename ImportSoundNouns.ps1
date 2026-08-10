<#
.SYNOPSIS
    Turns a numbered caption .txt into one Audacity label track per category.

.DESCRIPTION
    Bound to a hotkey by "Audacity annotation hotkey.ahk". The annotation site
    hands out files like "Hut R0_sound_nouns.txt" where every caption carries the
    category it belongs to:

        0.000<TAB>30.000<TAB>8 - Insect chirping
        0.200<TAB>0.933<TAB>1 - Male speech

    Audacity's own File > Import > Labels flattens all of that into one track.
    This builds a track per number instead - named "1", "2", "3", ... - with the
    "N - " prefix stripped from each caption, since the track already says it.

    Audacity has no scriptable "import labels from this path" (ImportLabels only
    opens a dialog, and Import2 rejects label files), so the tracks are built
    label by label over mod-script-pipe.

.PARAMETER Path
    Skip the file picker and read this .txt.

.PARAMETER Replace
    Answer the "replace existing label tracks?" prompt with yes, unattended.

.PARAMETER KeepExisting
    Answer it with no: leave existing label tracks alone and add alongside.

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

try {
    # Windows only lets the process that already owns the foreground window give
    # focus away. This one is a hidden powershell.exe started by the hotkey, so
    # SetForegroundWindow on its own is ignored and the dialog opens behind
    # Audacity. Attaching to the foreground window's input queue first makes the
    # two threads share a focus state, and the call is then honoured.
    if (-not ('AutoNyx.Foreground' -as [type])) {
        Add-Type -Namespace AutoNyx -Name Foreground -MemberDefinition @'
[DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
[DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr hWnd, IntPtr lpdwProcessId);
[DllImport("user32.dll")] private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool fAttach);
[DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr hWnd);
[DllImport("user32.dll")] private static extern bool BringWindowToTop(IntPtr hWnd);
[DllImport("user32.dll")] private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
[DllImport("user32.dll")] private static extern IntPtr GetLastActivePopup(IntPtr hWnd);
[DllImport("user32.dll")] private static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);
[DllImport("kernel32.dll")] private static extern uint GetCurrentThreadId();

public static void Force(IntPtr hWnd) {
    if (hWnd == IntPtr.Zero) { return; }
    IntPtr fore = GetForegroundWindow();
    uint us = GetCurrentThreadId();
    uint them = (fore == IntPtr.Zero) ? 0 : GetWindowThreadProcessId(fore, IntPtr.Zero);
    bool attached = false;
    if (them != 0 && them != us) { attached = AttachThreadInput(us, them, true); }
    try {
        ShowWindow(hWnd, 5);   // SW_SHOW
        BringWindowToTop(hWnd);
        SetForegroundWindow(hWnd);
    }
    finally {
        if (attached) { AttachThreadInput(us, them, false); }
    }
}

// The dialog itself, addressed through the stub form that owns it. A modal
// dialog is its owner's one ENABLEDPOPUP - ask for that rather than
// GetLastActivePopup, which reports the owner back while the dialog has never
// been active, which is precisely the situation here. No popup yet means no
// dialog yet, and raising the owner then is both harmless and what we want.
public static void ForcePopup(IntPtr owner) {
    if (owner == IntPtr.Zero) { return; }
    IntPtr popup = GetWindow(owner, 6);   // GW_ENABLEDPOPUP
    if (popup == IntPtr.Zero) { popup = GetLastActivePopup(owner); }
    Force(popup == IntPtr.Zero ? owner : popup);
}
'@
    }
}
catch {
    # No C# compiler, no focus fix: the dialogs open behind Audacity the way they
    # used to. That beats the import dying up here, above the try block that puts
    # failures on screen - launched hidden, it would vanish without a word.
}

$DialogTitle = 'Import sound nouns'

function Show-Problem {
    param([string]$Message)
    Invoke-WithOwner {
        param($owner)
        [System.Windows.Forms.MessageBox]::Show($owner, $Message, $DialogTitle,
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning) | Out-Null
    }
}

function Invoke-WithOwner {
    <#
        PowerShell is launched hidden by the hotkey, so a dialog has no natural
        owner and would open behind Audacity. A topmost stub form gives it one,
        and AutoNyx.Foreground makes the activation actually stick - Show() and
        Activate() alone are refused, since this process is not the one holding
        the foreground.
    #>
    param([Parameter(Mandatory)][scriptblock]$Action)

    $owner = New-Object System.Windows.Forms.Form
    $owner.TopMost = $true
    $owner.ShowInTaskbar = $false
    $owner.StartPosition = 'Manual'
    $owner.Location = New-Object System.Drawing.Point(-2000, -2000)
    $owner.Size = New-Object System.Drawing.Size(1, 1)
    $owner.Show(); $owner.Activate()

    $timer = $null
    if ('AutoNyx.Foreground' -as [type]) {
        [AutoNyx.Foreground]::Force($owner.Handle)

        # Raising the owner is the part that matters - once this process holds the
        # foreground, anything it opens afterwards may take focus freely. The timer
        # then catches the dialog itself. ShowDialog blocks, so a timer is the only
        # way to reach it once it is up; the modal loop pumps this thread's
        # messages, so it keeps ticking. A handful of ticks, then stop: past that
        # the user is the one deciding what has focus.
        $handle = $owner.Handle
        $ticks = 0
        $timer = New-Object System.Windows.Forms.Timer
        $timer.Interval = 120
        $timer.Add_Tick({
            $ticks++
            [AutoNyx.Foreground]::ForcePopup($handle)
            if ($ticks -ge 5) { $timer.Stop() }
        }.GetNewClosure())
        $timer.Start()
    }

    try { & $Action $owner }
    finally {
        if ($timer) { $timer.Stop(); $timer.Dispose() }
        $owner.Close(); $owner.Dispose()
    }
}

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
                       "Replace them with $($groups.Count) track(s) from $fileName?`n`n" +
                       "Yes  - delete them and import`n" +
                       "No   - leave this project alone"
            $answer = Invoke-WithOwner {
                param($owner)
                [System.Windows.Forms.MessageBox]::Show($owner, $message, $DialogTitle,
                    [System.Windows.Forms.MessageBoxButtons]::YesNo,
                    [System.Windows.Forms.MessageBoxIcon]::Question)
            }
            if ($answer -ne [System.Windows.Forms.DialogResult]::Yes) { exit 0 }
            $wipe = $true
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
    foreach ($group in $groups) {
        $commands.Add('NewLabelTrack:')
        # Select AND focus it. AddLabel follows the *focused* track, not the
        # selected one - selecting alone leaves focus whereever it was and the
        # labels land in somebody else's track.
        $commands.Add("SelectTracks: Track=$trackIndex TrackCount=1 Mode=Set")
        $commands.Add('SetTrack: Focused=1')
        $commands.Add("SetTrack: Name=`"$($group.Number)`"")

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

    foreach ($group in $groups) {
        $name = [string]$group.Number
        $got = @($after | Where-Object { $_.Track -eq $name } | Sort-Object Start, End)
        if ($got.Count -ne $group.Labels.Count) {
            $problems.Add("  track $name : expected $($group.Labels.Count) labels, found $($got.Count)")
            continue
        }
        for ($i = 0; $i -lt $got.Count; $i++) {
            $want = $group.Labels[$i]
            if ($got[$i].Text -ne $want.Text) {
                $problems.Add("  track $name label $($i + 1): expected '$($want.Text)', found '$($got[$i].Text)'")
            }
            elseif ([math]::Abs($got[$i].Start - $want.Start) -gt 0.001 -or
                    [math]::Abs($got[$i].End   - $want.End)   -gt 0.001) {
                $problems.Add(("  track {0} label {1}: expected {2:0.000}-{3:0.000}, found {4:0.000}-{5:0.000}" -f `
                    $name, ($i + 1), $want.Start, $want.End, $got[$i].Start, $got[$i].End))
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
