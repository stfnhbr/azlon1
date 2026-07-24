<#
    Talks to a running Audacity via mod-script-pipe.
    Dot-source this file; it defines functions only and runs nothing.

    Audacity creates two named pipes and serves ONE client at a time. Connecting
    to ToSrvPipe without also connecting FromSrvPipe wedges the server until the
    half-open handle is closed, so Invoke-AudacityCommands always does both and
    always disposes both.
#>

function Invoke-AudacityCommands {
    <#
    .SYNOPSIS
        Sends commands to Audacity over one pipe session; returns one response
        string per command, in order.
    #>
    param(
        [Parameter(Mandatory)] [string[]]$Commands,
        [int]$TimeoutMs = 5000
    )

    $enc = New-Object System.Text.UTF8Encoding($false)
    $out = New-Object System.IO.Pipes.NamedPipeClientStream('.', 'ToSrvPipe',   [System.IO.Pipes.PipeDirection]::Out)
    $in  = New-Object System.IO.Pipes.NamedPipeClientStream('.', 'FromSrvPipe', [System.IO.Pipes.PipeDirection]::In)

    try {
        try {
            $out.Connect($TimeoutMs)
            $in.Connect($TimeoutMs)
        } catch {
            throw "Could not reach Audacity's scripting pipe. Make sure Audacity is running, that Edit > Preferences > Modules > mod-script-pipe is set to Enabled, and that no other script is talking to it right now."
        }

        $w = New-Object System.IO.StreamWriter($out, $enc)
        $w.AutoFlush = $true
        $r = New-Object System.IO.StreamReader($in, $enc)

        $responses = @()
        foreach ($cmd in $Commands) {
            $w.Write("$cmd`r`n`0")

            $sb = New-Object System.Text.StringBuilder
            $status = $null
            while ($true) {
                $line = $r.ReadLine()
                if ($null -eq $line) { break }
                if ($line -match '^BatchCommand finished:\s*(.*)$') { $status = $Matches[1].Trim(); break }
                [void]$sb.AppendLine($line)
            }
            if ($status -ne 'OK') {
                throw "Audacity rejected '$cmd' (status: $status)"
            }
            $responses += $sb.ToString().Trim()
        }
        return , $responses
    }
    finally {
        $out.Dispose()
        $in.Dispose()
    }
}

function Get-AudacityAnnotations {
    <#
    .SYNOPSIS
        Reads every label from the running project, tagged with its label
        track's name, sorted by start time.
    .DESCRIPTION
        GetInfo Type=Labels numbers tracks 1..n across LABEL tracks only, while
        Type=Tracks lists every track. The label-kind tracks are filtered out in
        order so the two numberings line up.
    #>

    $responses = Invoke-AudacityCommands -Commands @(
        'GetInfo: Type=Tracks Format=JSON',
        'GetInfo: Type=Labels Format=JSON'
    )

    $labelTrackNames = @(($responses[0] | ConvertFrom-Json) |
        Where-Object { $_.kind -eq 'label' } |
        ForEach-Object { $_.name })

    $rows = New-Object System.Collections.Generic.List[object]
    foreach ($entry in ($responses[1] | ConvertFrom-Json)) {
        # entry = [ trackIndex, [ [start, end, text], ... ] ]
        $idx    = [int]$entry[0]
        $name   = if ($idx -ge 1 -and $idx -le $labelTrackNames.Count) { $labelTrackNames[$idx - 1] } else { "Track $idx" }
        foreach ($lab in $entry[1]) {
            $rows.Add([pscustomobject]@{
                Text        = [string]$lab[2]
                Start       = [double]$lab[0]
                End         = [double]$lab[1]
                Track       = $name
                TrackNumber = $idx
            })
        }
    }

    # Stable within a timestamp: ties fall back to the track's on-screen order.
    return , @($rows | Sort-Object Start, TrackNumber, End)
}

function Get-AudacityProjectContext {
    <#
    .SYNOPSIS
        Best-effort project name and folder for the running Audacity window.
    .DESCRIPTION
        The scripting API exposes no project path, so the window title supplies
        the name and the recent-files list supplies the folder it was opened from.
    #>

    $proc = Get-Process -Name Audacity -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowTitle } | Select-Object -First 1

    $name = if ($proc) { $proc.MainWindowTitle } else { '' }
    $name = ($name -replace '\s*[-–]\s*Audacity\s*$', '').Trim().TrimEnd('*').Trim()
    if (-not $name) { $name = 'Annotations' }

    $dir = $null
    $cfgPath = Join-Path $env:APPDATA 'audacity\audacity.cfg'
    if (Test-Path -LiteralPath $cfgPath) {
        $cfg = Get-Content -LiteralPath $cfgPath -Raw
        if ($cfg -match '(?s)\[RecentFiles\](.*?)(\r?\n\[|$)') {
            foreach ($line in ($Matches[1] -split '\r?\n')) {
                if ($line -notmatch '^\s*file\d+=(.+)$') { continue }
                $path = $Matches[1].Replace('\\', '\').Trim()
                if ([System.IO.Path]::GetFileNameWithoutExtension($path) -ne $name) { continue }
                $candidate = [System.IO.Path]::GetDirectoryName($path)
                if (Test-Path -LiteralPath $candidate) { $dir = $candidate; break }
            }
        }
    }
    if (-not $dir) { $dir = [Environment]::GetFolderPath('Desktop') }

    return [pscustomobject]@{ Name = $name; Directory = $dir }
}
