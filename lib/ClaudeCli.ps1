<#
    Runs one prompt through the Claude Code CLI in headless mode and hands back
    what the model said.

    The CLI rather than the API on purpose: it is already installed and already
    signed in, so there is no key to store on this machine and nothing extra to
    pay for. The cost is a few seconds of start-up per call.

    Dot-source this file; it defines functions only and runs nothing.
#>

function Get-ClaudeExecutable {
    <#
    .SYNOPSIS
        Full path to claude.exe, or $null.
    .DESCRIPTION
        npm puts three shims on PATH - claude, claude.cmd and claude.ps1 - and
        Get-Command reports the .ps1 first. None of them can be started with
        UseShellExecute off, which is what redirecting stdin needs, so the real
        executable behind them is what we want: it lives beside the package the
        shims call into.
    #>
    $direct = @(Get-Command claude -All -ErrorAction SilentlyContinue |
                Where-Object { $_.Source -and $_.Source.ToLower().EndsWith('.exe') })
    if ($direct.Count -gt 0) { return $direct[0].Source }

    $shim = Get-Command claude -ErrorAction SilentlyContinue
    if (-not $shim -or -not $shim.Source) { return $null }

    $candidate = Join-Path (Split-Path -Parent $shim.Source) `
                           'node_modules\@anthropic-ai\claude-code\bin\claude.exe'
    if (Test-Path -LiteralPath $candidate) { return (Resolve-Path -LiteralPath $candidate).ProviderPath }

    return $null
}

function Invoke-ClaudeJson {
    <#
    .SYNOPSIS
        Sends one prompt to Claude and returns the reply text.
    .DESCRIPTION
        The prompt goes in over stdin rather than as an argument: a few hundred
        caption rows would run past the command-line length limit. It is written
        through the raw stream with an explicit UTF-8 writer, because
        ProcessStartInfo has no StandardInputEncoding on .NET Framework and
        PowerShell 5.1 would otherwise send ASCII and flatten every accented
        character in the annotation text.

        --output-format json wraps the reply in a result envelope, which is
        where a refusal or an internal error shows up as something other than a
        non-zero exit code.
    .PARAMETER Tick
        Called every poll while waiting. The caller passes DoEvents through this
        so a progress window keeps painting - nothing else pumps messages while
        this blocks.
    .PARAMETER TimeoutMs
        Kill the CLI after this long. A hidden hotkey process that wedges is
        invisible, so there is always a ceiling.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Prompt,
        [string]$Model,
        [int]$TimeoutMs = 300000,
        [scriptblock]$Tick
    )

    $exe = Get-ClaudeExecutable
    if (-not $exe) {
        throw ("Could not find the Claude Code CLI.`n`n" +
               "Install it with:  npm install -g @anthropic-ai/claude-code`n" +
               "then run 'claude' once to sign in.")
    }

    # --allowed-tools with an empty list: this is a text transformation, so the
    # model has no business touching the disk, and a tool call in a headless run
    # only ends in a permission refusal anyway.
    $arguments = @('-p', '--output-format', 'json', '--allowed-tools', '""')
    if ($Model) { $arguments += @('--model', $Model) }

    $psi = New-Object System.Diagnostics.ProcessStartInfo
    $psi.FileName               = $exe
    $psi.Arguments              = ($arguments -join ' ')
    $psi.UseShellExecute        = $false
    $psi.CreateNoWindow         = $true
    $psi.RedirectStandardInput  = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError  = $true
    $psi.StandardOutputEncoding = New-Object System.Text.UTF8Encoding $false
    $psi.StandardErrorEncoding  = New-Object System.Text.UTF8Encoding $false
    $psi.WorkingDirectory       = $env:TEMP

    $proc = New-Object System.Diagnostics.Process
    $proc.StartInfo = $psi

    try {
        [void]$proc.Start()

        $writer = New-Object System.IO.StreamWriter($proc.StandardInput.BaseStream,
                                                    (New-Object System.Text.UTF8Encoding $false))
        $writer.Write($Prompt)
        $writer.Flush()
        $writer.Close()

        # Read both pipes as they fill. Waiting for exit first and reading after
        # deadlocks the moment the reply outgrows the pipe buffer.
        $stdout = $proc.StandardOutput.ReadToEndAsync()
        $stderr = $proc.StandardError.ReadToEndAsync()

        $waited = 0
        while (-not $proc.WaitForExit(150)) {
            $waited += 150
            if ($Tick) { & $Tick }
            if ($waited -ge $TimeoutMs) {
                try { $proc.Kill() } catch { }
                throw "Claude did not answer within $([int]($TimeoutMs / 1000)) seconds."
            }
        }

        $out = $stdout.Result
        $err = $stderr.Result

        if ($proc.ExitCode -ne 0) {
            $detail = if ($err.Trim()) { $err.Trim() } else { $out.Trim() }
            throw "The Claude CLI exited with code $($proc.ExitCode).`n`n$detail"
        }

        if (-not $out.Trim()) { throw "The Claude CLI returned nothing.`n`n$($err.Trim())" }

        $envelope = $null
        try { $envelope = $out | ConvertFrom-Json }
        catch { throw "Could not read the Claude CLI's reply as JSON:`n`n$($out.Trim())" }

        if ($envelope.is_error) {
            throw "Claude reported an error:`n`n$($envelope.result)"
        }
        if ($null -eq $envelope.result) {
            throw "The Claude CLI's reply carried no result:`n`n$($out.Trim())"
        }

        return [string]$envelope.result
    }
    finally {
        $proc.Dispose()
    }
}

function ConvertFrom-ClaudeJsonBlock {
    <#
    .SYNOPSIS
        Parses JSON out of a model reply, fenced or not.
    .DESCRIPTION
        Asked for JSON, the model usually returns exactly that, and sometimes
        returns it inside a ```json fence or with a sentence in front. Rather
        than insist, take the outermost [...] or {...} and parse that.

        Assign before wrapping at the call site: Windows PowerShell 5.1 pushes a
        JSON array down the pipeline as ONE object, so @(... | ConvertFrom-Json)
        collapses every element into a single item.
    #>
    param([Parameter(Mandatory)][string]$Text)

    $body = $Text.Trim()
    if ($body -match '(?s)```(?:json)?\s*(.+?)\s*```') { $body = $Matches[1].Trim() }

    if ($body -notmatch '(?s)^\s*[\[\{]') {
        $open = $body.IndexOfAny([char[]]@('[', '{'))
        $close = [Math]::Max($body.LastIndexOf(']'), $body.LastIndexOf('}'))
        if ($open -lt 0 -or $close -le $open) {
            throw "Claude's reply is not JSON:`n`n$($Text.Trim())"
        }
        $body = $body.Substring($open, $close - $open + 1)
    }

    try { return $body | ConvertFrom-Json }
    catch { throw "Could not read Claude's reply as JSON:`n`n$($Text.Trim())" }
}
