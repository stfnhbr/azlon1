<#
    Which Audacity is running, and which project file a hotkey press means.

    Dot-source this file; it defines functions only and runs nothing.

    Two Audacitys now live side by side here:

        Audacity.exe   3.7.9  C:\Program Files\Audacity        - has the pipe
        Audacity4.exe  4.0    C:\Program Files\Audacity 4\bin  - has no pipe

    3.7.9 is driven over mod-script-pipe as it always was. 4.0 has no scripting
    interface of any kind, so it is reached through its project file instead,
    and the job of this file is to work out which file that is.

    That question is harder than it sounds, and getting it wrong is expensive:
    see the note about the pipe firing at whichever project is frontmost. The
    rule here is the opposite one - never guess silently. One open project is
    used without asking; several means asking.
#>

if (-not ([System.Management.Automation.PSTypeName]'AutoNyx.AuWindows').Type) {
Add-Type -TypeDefinition @'
using System;
using System.Text;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace AutoNyx
{
    public class AuWindow
    {
        public IntPtr Handle;
        public int ProcessId;
        public string Title;
    }

    public static class AuWindows
    {
        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        static extern bool EnumWindows(EnumProc cb, IntPtr p);
        [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "GetWindowTextW")]
        static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
        [DllImport("user32.dll")]
        static extern bool IsWindowVisible(IntPtr h);
        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(IntPtr h, out int pid);
        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        delegate bool EnumProc(IntPtr h, IntPtr p);

        // Qt puts a swarm of hidden helper windows on the process; only visible
        // ones with a real title are project windows.
        public static List<AuWindow> Visible(int[] pids)
        {
            List<AuWindow> found = new List<AuWindow>();
            EnumWindows(delegate(IntPtr h, IntPtr p)
            {
                if (!IsWindowVisible(h)) return true;
                int pid; GetWindowThreadProcessId(h, out pid);
                bool wanted = false;
                foreach (int want in pids) if (want == pid) { wanted = true; break; }
                if (!wanted) return true;
                StringBuilder sb = new StringBuilder(1024);
                GetWindowText(h, sb, sb.Capacity);
                string t = sb.ToString();
                if (t.Length < 2 || t == "_q_titlebar") return true;
                AuWindow w = new AuWindow();
                w.Handle = h; w.ProcessId = pid; w.Title = t;
                found.Add(w);
                return true;
            }, IntPtr.Zero);
            return found;
        }

        public static int ForegroundProcessId()
        {
            IntPtr h = GetForegroundWindow();
            if (h == IntPtr.Zero) return 0;
            int pid; GetWindowThreadProcessId(h, out pid);
            return pid;
        }
    }
}
'@
}

# ---------------------------------------------------------------------------

<#
.SYNOPSIS
    What Audacity is running right now, and which one is in front.
.DESCRIPTION
    Version is 3 or 4. The .Focused member is the one whose window currently
    holds the foreground, which is what a hotkey pressed inside Audacity means -
    and with both versions open at once it is the only honest answer.
#>
function Get-AudacityAppState {
    [CmdletBinding()] param()

    $apps = @()
    foreach ($spec in @(
        @{ Version = 3; ProcessName = 'Audacity'  },
        @{ Version = 4; ProcessName = 'Audacity4' }
    )) {
        $procs = @(Get-Process -Name $spec.ProcessName -ErrorAction SilentlyContinue)
        if ($procs.Count -eq 0) { continue }
        $windows = @()
        if ($procs.Count -gt 0) {
            $windows = @([AutoNyx.AuWindows]::Visible([int[]]@($procs | ForEach-Object { $_.Id })))
        }
        $apps += [pscustomobject]@{
            Version     = $spec.Version
            ProcessName = $spec.ProcessName
            ProcessIds  = @($procs | ForEach-Object { $_.Id })
            Windows     = $windows
        }
    }

    $foregroundPid = [AutoNyx.AuWindows]::ForegroundProcessId()
    $focused = $null
    foreach ($app in $apps) { if ($app.ProcessIds -contains $foregroundPid) { $focused = $app.Version } }

    return [pscustomobject]@{
        Apps      = $apps
        Focused   = $focused
        Running3  = [bool]($apps | Where-Object { $_.Version -eq 3 })
        Running4  = [bool]($apps | Where-Object { $_.Version -eq 4 })
    }
}

<#
.SYNOPSIS
    Decides which Audacity a run is aimed at: 3 or 4.
.DESCRIPTION
    In order: what the caller asked for, then whichever Audacity is in front,
    then the only one running. Throws rather than picking a side when both are
    open and neither is focused, because the two routes write to different
    places and a wrong guess is not a small mistake.
#>
function Resolve-AudacityVersion {
    [CmdletBinding()]
    param(
        [ValidateSet('3', '4', 'Auto')] [string]$Version = 'Auto',
        $State
    )

    if ($Version -ne 'Auto') { return [int]$Version }
    if (-not $State) { $State = Get-AudacityAppState }

    if ($State.Focused) { return $State.Focused }
    if ($State.Running4 -and -not $State.Running3) { return 4 }
    if ($State.Running3 -and -not $State.Running4) { return 3 }
    if (-not $State.Running3 -and -not $State.Running4) {
        throw "Audacity is not running. Open the project first, then press the hotkey again."
    }
    throw ("Both Audacity 3.7.9 and Audacity 4.0 are open, and neither is the window in front, " +
           "so there is no telling which one you meant.`n`n" +
           "Click the Audacity you want and press the hotkey again.")
}

<#
.SYNOPSIS
    True when something else holds the file open - which is how a project that
    Audacity has loaded announces itself.
.DESCRIPTION
    SQLite opens its database with read/write sharing, so asking for the file
    with no sharing at all fails for exactly as long as Audacity has it. Note
    this is the only reliable test: a NAME.aup3-wal beside the project proves
    nothing, since one is left behind by any crash and by any read-only open.
#>
function Test-AudacityFileOpen {
    [CmdletBinding()] param([Parameter(Mandatory)] [string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) { return $false }
    try {
        $handle = [System.IO.File]::Open($Path, 'Open', 'ReadWrite', 'None')
        $handle.Close(); $handle.Dispose()
        return $false
    } catch { return $true }
}

<#
.SYNOPSIS
    The project files Audacity 4.0 currently has open.
.DESCRIPTION
    4.0 records the projects it has loaded under [ActiveProjects] in
    Audacity4.ini, and parks not-yet-saved ones in its SessionData folder as
    NAME.aup4unsaved - an ordinary project database under another extension,
    which is why unsaved work can be read at all.

    That list is a history rather than a live register: it keeps entries for
    projects closed long ago. What makes an entry current is that Audacity still
    holds the file open, so every candidate is tested and only the held ones
    come back.
#>
function Get-Audacity4OpenProjects {
    [CmdletBinding()] param()

    $iniPath = Join-Path $env:APPDATA 'audacity\Audacity4.ini'
    $candidates = New-Object System.Collections.Generic.List[string]
    $sessionDir = Join-Path $env:LOCALAPPDATA 'Audacity\Audacity4\SessionData'

    if (Test-Path -LiteralPath $iniPath) {
        $ini = Get-Content -LiteralPath $iniPath -Raw
        if ($ini -match '(?s)\[ActiveProjects\](.*?)(\r?\n\[|$)') {
            foreach ($line in ($Matches[1] -split '\r?\n')) {
                if ($line -notmatch '^\s*\d+\s*=\s*(.+)$') { continue }
                # The ini escapes backslashes, and quotes a value with spaces.
                $entry = $Matches[1].Trim().Trim('"').Replace('\\', '\')
                if ($entry) { $candidates.Add($entry) }
            }
        }
        # TempDir is where the unsaved ones live; trust the ini over the guess.
        if ($ini -match '(?m)^\s*TempDir\s*=\s*(.+)$') {
            $fromIni = $Matches[1].Trim().Trim('"').Replace('\\', '\').Replace('/', '\')
            if ($fromIni) { $sessionDir = $fromIni }
        }
    }

    if (Test-Path -LiteralPath $sessionDir) {
        Get-ChildItem -LiteralPath $sessionDir -Filter *.aup4unsaved -File -ErrorAction SilentlyContinue |
            ForEach-Object { $candidates.Add($_.FullName) }
    }

    $seen = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    $open = New-Object System.Collections.Generic.List[object]
    foreach ($path in $candidates) {
        if (-not $seen.Add($path)) { continue }
        if (-not (Test-Path -LiteralPath $path)) { continue }
        if (-not (Test-AudacityFileOpen -Path $path)) { continue }
        $item = Get-Item -LiteralPath $path
        $open.Add([pscustomobject]@{
            Path     = $path
            Saved    = ($item.Extension -ne '.aup4unsaved')
            Modified = $item.LastWriteTime
            Name     = Get-AudacityProjectDisplayName -Path $path
        })
    }

    return , @($open | Sort-Object Modified -Descending)
}

<#
.SYNOPSIS
    The folder Audacity 4.0 last opened or saved a project in.
.DESCRIPTION
    Somewhere to point a file dialog. 4.0 keeps it under [project] as
    paths\lastprojects; the Desktop is the fallback, as elsewhere in the
    toolchain.
#>
function Get-Audacity4DefaultFolder {
    [CmdletBinding()] param()

    $iniPath = Join-Path $env:APPDATA 'audacity\Audacity4.ini'
    if (Test-Path -LiteralPath $iniPath) {
        $ini = Get-Content -LiteralPath $iniPath -Raw
        if ($ini -match '(?m)^\s*paths\\lastprojects\s*=\s*(.+)$') {
            $folder = $Matches[1].Trim().Trim('"').Replace('\\', '\').Replace('/', '\')
            if ($folder -and (Test-Path -LiteralPath $folder)) { return $folder }
        }
    }
    return [Environment]::GetFolderPath('Desktop')
}

<#
.SYNOPSIS
    A name for a project file worth showing a human.
.DESCRIPTION
    A saved project is named by its file. An unsaved one is not: its file is
    called "New Project 2026-09-09 13-21-02 N-1.aup4unsaved" no matter what the
    window says. Audacity titles that window after the first track, so the first
    track name is read out of the document and used, which is what makes the
    picker line up with what is on screen.
#>
function Get-AudacityProjectDisplayName {
    [CmdletBinding()] param([Parameter(Mandatory)] [string]$Path)

    if ([System.IO.Path]::GetExtension($Path) -ne '.aup4unsaved') {
        return [System.IO.Path]::GetFileNameWithoutExtension($Path)
    }
    try {
        $document = Read-AudacityProjectFile -Path $Path
        if ($document.TrackNames.Count -gt 0) { return $document.TrackNames[0] }
    } catch { }
    return 'Unsaved project'
}

<#
.SYNOPSIS
    Settles on one Audacity 4.0 project file, asking when there is a choice.
#>
function Select-Audacity4ProjectFile {
    [CmdletBinding()]
    param([string]$Title = 'Audacity 4 project')

    $open = @(Get-Audacity4OpenProjects)

    if ($open.Count -eq 1) { return $open[0] }
    if ($open.Count -eq 0) {
        throw ("Audacity 4.0 does not appear to have a project open.`n`n" +
               "If it does, save it once - a project that has never been saved and has no " +
               "unsaved changes leaves nothing on disk to read.")
    }

    # More than one. Ask, rather than reach for the frontmost the way the pipe
    # does: that habit is what once put labels in the wrong project.
    $choice = Show-AudacityProjectChoice -Projects $open -Title $Title
    if (-not $choice) { return $null }
    return $choice
}

<#
.SYNOPSIS
    A small list dialog for choosing between open projects.
#>
function Show-AudacityProjectChoice {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)] $Projects,
        [string]$Title = 'Which project?'
    )

    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.StartPosition = 'CenterScreen'
    $form.FormBorderStyle = 'FixedDialog'
    $form.MinimizeBox = $false; $form.MaximizeBox = $false
    $form.ClientSize = New-Object System.Drawing.Size(520, 290)

    $label = New-Object System.Windows.Forms.Label
    $label.Text = "Audacity 4.0 has more than one project open. Which one?"
    $label.SetBounds(12, 12, 496, 20)
    $form.Controls.Add($label)

    $list = New-Object System.Windows.Forms.ListBox
    $list.SetBounds(12, 38, 496, 190)
    foreach ($project in $Projects) {
        $when = $project.Modified.ToString('HH:mm')
        $kind = if ($project.Saved) { $project.Path } else { 'not saved yet' }
        [void]$list.Items.Add(("{0}    (changed {1})    {2}" -f $project.Name, $when, $kind))
    }
    $list.SelectedIndex = 0
    $form.Controls.Add($list)

    $ok = New-Object System.Windows.Forms.Button
    $ok.Text = 'OK'; $ok.SetBounds(332, 240, 84, 28)
    $ok.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Controls.Add($ok); $form.AcceptButton = $ok

    $cancel = New-Object System.Windows.Forms.Button
    $cancel.Text = 'Cancel'; $cancel.SetBounds(424, 240, 84, 28)
    $cancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($cancel); $form.CancelButton = $cancel

    # Invoke-WithOwner comes from lib\Dialogs.ps1: PowerShell is started hidden
    # by the hotkey, so a dialog with no owner opens behind Audacity unseen.
    $result = if (Get-Command Invoke-WithOwner -ErrorAction SilentlyContinue) {
        Invoke-WithOwner { param($owner) $form.ShowDialog($owner) }
    } else { $form.ShowDialog() }

    $index = $list.SelectedIndex
    $form.Dispose()
    if ($result -ne [System.Windows.Forms.DialogResult]::OK -or $index -lt 0) { return $null }
    return $Projects[$index]
}
