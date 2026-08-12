<#
    Dialogs for the hotkey-launched scripts, which run as a hidden
    powershell.exe and therefore have no window to own anything they open.

    Dot-source this file, then set $DialogTitle to whatever the message boxes
    should be titled. It defines functions and the AutoNyx.Foreground type and
    runs nothing else.

    Requires System.Windows.Forms and System.Drawing; the caller adds them.
#>

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
    # used to. That beats the script dying up here, above the try block that puts
    # failures on screen - launched hidden, it would vanish without a word.
}

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

function New-ProgressWindow {
    <#
    .SYNOPSIS
        A small always-on-top status window. Returns an object with SetText()
        and Close(); the caller must Close() it.
    .DESCRIPTION
        Not a dialog and not a prompt - it has no buttons and blocks nothing.
        The hotkey starts this process hidden, so without it a run that spends
        half a minute waiting on Claude looks like a keypress that did nothing.

        Nothing pumps messages while the caller works, so the caller passes
        Application::DoEvents as the Tick of whatever it is waiting on. Painting
        stops without it, which is why SetText calls Refresh rather than trusting
        the next paint to arrive.
    #>
    param(
        [Parameter(Mandatory)][string]$Text,
        [string]$Title = $DialogTitle
    )

    $form = New-Object System.Windows.Forms.Form
    $form.Text = $Title
    $form.FormBorderStyle = 'FixedToolWindow'
    $form.StartPosition = 'CenterScreen'
    $form.TopMost = $true
    $form.ShowInTaskbar = $false
    $form.ControlBox = $false
    $form.Size = New-Object System.Drawing.Size(420, 96)

    $label = New-Object System.Windows.Forms.Label
    $label.AutoSize = $false
    $label.Dock = 'Fill'
    $label.TextAlign = 'MiddleCenter'
    $label.Text = $Text
    $form.Controls.Add($label)

    $form.Show()
    if ('AutoNyx.Foreground' -as [type]) { [AutoNyx.Foreground]::Force($form.Handle) }
    [System.Windows.Forms.Application]::DoEvents()

    return [pscustomobject]@{
        Form  = $form
        Label = $label
    } | Add-Member -PassThru -MemberType ScriptMethod -Name SetText -Value {
        param([string]$Value)
        $this.Label.Text = $Value
        $this.Label.Refresh()
        [System.Windows.Forms.Application]::DoEvents()
    } | Add-Member -PassThru -MemberType ScriptMethod -Name Close -Value {
        if ($this.Form) { $this.Form.Close(); $this.Form.Dispose(); $this.Form = $null }
    }
}
