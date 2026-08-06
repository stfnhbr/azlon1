<#
    Reads the numbered caption files the annotation site hands out, e.g.
    "Hut R0_sound_nouns.txt":

        0.000<TAB>30.000<TAB>8 - Insect chirping
        0.200<TAB>0.933<TAB>1 - Male speech

    Tab separated, UTF-8, and the number before the dash is the category the
    caption belongs to - the thing Audacity's own label import throws away by
    flattening every line into one track.

    Dot-source this file; it defines functions only and runs nothing.
#>

# Start, End, then everything else. Audacity writes label times with a dot
# regardless of locale, so the pattern does too.
$script:SoundNounLine = '^(?<start>-?[0-9]+(\.[0-9]+)?)\t(?<end>-?[0-9]+(\.[0-9]+)?)\t(?<rest>.*)$'

# "8 - Insect chirping". The separator is an en-dash (U+2013) in the files seen
# so far; hyphen and em-dash are accepted too in case the exporter changes its
# mind. Anchored and non-greedy at the front, so a dash inside the caption
# itself stays part of the caption.
#
# The dashes are written as \u escapes rather than literally: Windows
# PowerShell 5.1 reads a .ps1 without a BOM as ANSI, which would corrupt a
# literal en-dash here and stop every line from matching. Keeping this file
# pure ASCII sidesteps that entirely.
$script:SoundNounPrefix = '^(?<num>[0-9]+)\s*[\u2013\u2014-]\s*(?<text>.+)$'

function Read-SoundNounFile {
    <#
    .SYNOPSIS
        Parses a sound-nouns file into one group per category number.
    .DESCRIPTION
        Returns an array of objects:

            Number : [int]    the category, ascending
            Labels : [array]  Start, End, Text - sorted by Start then End

        The caption's "N - " prefix is stripped; the group's Number carries it.

        Overlapping labels within a category are left alone - they are real
        (one file has a category running 5.533-14.533 and 13.267-20.600) and
        Audacity label tracks handle them.
    .PARAMETER Path
        The .txt to read.
    #>
    [CmdletBinding()]
    param([Parameter(Mandatory)][string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Cannot find the caption file:`n$Path"
    }

    # Read as UTF-8 explicitly. Get-Content on Windows PowerShell 5.1 falls back
    # to the ANSI codepage, which turns the en-dash separator into two bytes of
    # mojibake and nothing matches afterwards.
    $text  = [System.IO.File]::ReadAllText($Path, [System.Text.Encoding]::UTF8)
    $lines = $text -split "`r?`n"

    $labels     = New-Object System.Collections.Generic.List[object]
    $unnumbered = New-Object System.Collections.Generic.List[string]
    $malformed  = New-Object System.Collections.Generic.List[string]

    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        if ($line.Trim() -eq '') { continue }
        if ($line -match '^\\') { continue }   # frequency range continuation

        if ($line -notmatch $script:SoundNounLine) {
            $malformed.Add("  line $($i + 1): $line")
            continue
        }
        # Stash these now - the next -match overwrites $Matches.
        $start = $Matches['start']
        $end   = $Matches['end']
        $rest  = $Matches['rest'].Trim()

        if ($rest -notmatch $script:SoundNounPrefix) {
            $unnumbered.Add("  line $($i + 1): $rest")
            continue
        }

        $labels.Add([pscustomobject]@{
            Number = [int]$Matches['num']
            Text   = $Matches['text'].Trim()
            Start  = [double]::Parse($start, [System.Globalization.CultureInfo]::InvariantCulture)
            End    = [double]::Parse($end,   [System.Globalization.CultureInfo]::InvariantCulture)
        })
    }

    # Stop rather than guess. These files are machine-generated and uniform, so
    # a line that does not fit means something changed upstream - and filing a
    # caption under the wrong category silently is worse than not importing.
    if ($malformed.Count -gt 0) {
        throw ("$([System.IO.Path]::GetFileName($Path)): $($malformed.Count) line(s) are not " +
               "'start<TAB>end<TAB>caption':`n" + ($malformed -join "`n"))
    }
    if ($unnumbered.Count -gt 0) {
        throw ("$([System.IO.Path]::GetFileName($Path)): $($unnumbered.Count) caption(s) have no " +
               "'N - ' category prefix, so there is no track to put them on:`n" +
               ($unnumbered -join "`n"))
    }
    if ($labels.Count -eq 0) {
        throw "$([System.IO.Path]::GetFileName($Path)) holds no captions."
    }

    $groups = $labels |
        Group-Object Number |
        Sort-Object { [int]$_.Name } |
        ForEach-Object {
            [pscustomobject]@{
                Number = [int]$_.Name
                Labels = @($_.Group | Sort-Object Start, End)
            }
        }

    return , @($groups)
}

function Find-SoundNounFile {
    <#
    .SYNOPSIS
        Best guess at the caption file belonging to a project, or $null.
    .DESCRIPTION
        The downloads arrive percent-encoded and not always in the project's
        casing - "HUt%20R0_sound_nouns.txt" against a project called "Hut" - so
        the name is unescaped and compared case-insensitively.

        Any .txt starting with the project name counts, since the site names
        them inconsistently ("Cast20-%20Annotations.txt" alongside
        "HUt%20R0_sound_nouns.txt"); the sound-nouns ones are preferred, then
        the most recent.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Directory,
        [Parameter(Mandatory)][string]$ProjectName
    )

    if (-not (Test-Path -LiteralPath $Directory)) { return $null }

    $candidates = @(Get-ChildItem -LiteralPath $Directory -Filter '*.txt' -File -ErrorAction SilentlyContinue)
    if ($candidates.Count -eq 0) { return $null }

    $match = $candidates | Where-Object {
        $_.Name.StartsWith($ProjectName, [System.StringComparison]::OrdinalIgnoreCase)
    } | Sort-Object @{ Expression = { $_.Name -like '*sound_nouns*' }; Descending = $true },
                    @{ Expression = 'LastWriteTime';                  Descending = $true } |
        Select-Object -First 1

    if ($match) { return $match.FullName }
    return $null
}
