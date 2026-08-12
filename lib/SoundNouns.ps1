<#
    Reads the numbered caption files the annotation site hands out, e.g.
    "Hut R0_sound_nouns.txt":

        0.000<TAB>30.000<TAB>8 - Insect chirping
        0.200<TAB>0.933<TAB>1 - Male speech

    Tab separated, UTF-8, and the number before the dash is the category the
    caption belongs to - the thing Audacity's own label import throws away by
    flattening every line into one track.

    A fourth field may carry the category's track name, straight from the
    workbook the file was made from:

        0.200<TAB>0.933<TAB>1 - Male speech<TAB>Male Speech

    It is optional. Three-field files are the older shape and still read.

    Dot-source this file; it defines functions only and runs nothing.
#>

# Start, End, the caption, and optionally the track name. Audacity writes label
# times with a dot regardless of locale, so the pattern does too.
#
# The caption is [^\t]* rather than .* on purpose: .* would swallow the tab and
# the track name behind it, and since . matches a tab the name would then end up
# inside the caption text. A line carrying a fifth field fails to match as a
# result and is reported as malformed - which is the intent, see the throw below.
$script:SoundNounLine = '^(?<start>-?[0-9]+(\.[0-9]+)?)\t(?<end>-?[0-9]+(\.[0-9]+)?)\t(?<rest>[^\t]*)(\t(?<name>[^\t]*))?$'

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
            Name   : [string] the track name, '' when the file has three fields
            Labels : [array]  Start, End, Text, Line - sorted by Start then End

        The caption's "N - " prefix is stripped; the group's Number carries it.

        Rows are grouped by their category number and the groups sorted
        ascending, so neither the order of the lines nor a gap in the numbering
        matters - the files are written straight out of a workbook that is
        sorted by track name, not by number or by time.

        Every row of a category must agree on the track name, or this throws:
        naming the track after whichever row happened to come first would be
        guessing. A name shared by two categories only warns - the tracks still
        end up distinctly named, since the number leads.

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
        # Stash these now - the next -match overwrites $Matches. The name comes
        # from a group that need not have taken part, so cast before trimming:
        # a three-field line leaves it $null.
        $start = $Matches['start']
        $end   = $Matches['end']
        $rest  = $Matches['rest'].Trim()
        $name  = ([string]$Matches['name']).Trim()

        if ($rest -notmatch $script:SoundNounPrefix) {
            $unnumbered.Add("  line $($i + 1): $rest")
            continue
        }

        $labels.Add([pscustomobject]@{
            Number = [int]$Matches['num']
            Text   = $Matches['text'].Trim()
            Name   = $name
            Line   = $i + 1
            Start  = [double]::Parse($start, [System.Globalization.CultureInfo]::InvariantCulture)
            End    = [double]::Parse($end,   [System.Globalization.CultureInfo]::InvariantCulture)
        })
    }

    # Stop rather than guess. These files are machine-generated and uniform, so
    # a line that does not fit means something changed upstream - and filing a
    # caption under the wrong category silently is worse than not importing.
    if ($malformed.Count -gt 0) {
        throw ("$([System.IO.Path]::GetFileName($Path)): $($malformed.Count) line(s) are not " +
               "'start<TAB>end<TAB>caption' with an optional '<TAB>track name':`n" +
               ($malformed -join "`n"))
    }
    if ($unnumbered.Count -gt 0) {
        throw ("$([System.IO.Path]::GetFileName($Path)): $($unnumbered.Count) caption(s) have no " +
               "'N - ' category prefix, so there is no track to put them on:`n" +
               ($unnumbered -join "`n"))
    }
    if ($labels.Count -eq 0) {
        throw "$([System.IO.Path]::GetFileName($Path)) holds no captions."
    }

    # The names each category was given, in the order they first appeared, with
    # the line that introduced them. Ordinal so that "Male speech" and "Male
    # Speech" count as two names: which casing is the real one is not something
    # to decide here. A blank name is not a name - the field simply was not
    # supplied on that line - so it never conflicts with a filled one.
    $namesByNumber = @{}
    foreach ($label in $labels) {
        if ($label.Name -eq '') { continue }
        if (-not $namesByNumber.ContainsKey($label.Number)) {
            $namesByNumber[$label.Number] =
                New-Object System.Collections.Specialized.OrderedDictionary ([System.StringComparer]::Ordinal)
        }
        $seen = $namesByNumber[$label.Number]
        if (-not $seen.Contains($label.Name)) { $seen[$label.Name] = $label.Line }
    }

    # Stop rather than guess, again. One category means one track, so two names
    # under one number leaves nothing to call it.
    $conflicts = New-Object System.Collections.Generic.List[string]
    foreach ($number in ($namesByNumber.Keys | Sort-Object)) {
        $seen = $namesByNumber[$number]
        if ($seen.Count -le 1) { continue }
        $detail = @(foreach ($key in $seen.Keys) { "line $($seen[$key]) '$key'" }) -join ', '
        $conflicts.Add("  category ${number}: $detail")
    }
    if ($conflicts.Count -gt 0) {
        throw ("$([System.IO.Path]::GetFileName($Path)): $($conflicts.Count) category/categories carry " +
               "more than one track name, so there is no name to give the track:`n" +
               ($conflicts -join "`n"))
    }

    # The other direction only warns. Two categories called the same thing still
    # produce distinct Audacity tracks, because the number leads the name - but
    # it says the workbook has one track split across two numbers, which is
    # worth hearing about. Warning, not a dialog: it is not a reason to stop.
    $numbersByName = New-Object System.Collections.Specialized.OrderedDictionary ([System.StringComparer]::Ordinal)
    foreach ($number in ($namesByNumber.Keys | Sort-Object)) {
        foreach ($key in $namesByNumber[$number].Keys) {
            if (-not $numbersByName.Contains($key)) {
                $numbersByName[$key] = New-Object System.Collections.Generic.List[int]
            }
            $numbersByName[$key].Add([int]$number)
        }
    }
    foreach ($key in $numbersByName.Keys) {
        if ($numbersByName[$key].Count -le 1) { continue }
        Write-Warning ("$([System.IO.Path]::GetFileName($Path)): '$key' is the track name for " +
                       "categories $($numbersByName[$key] -join ', ').")
    }

    $groups = $labels |
        Group-Object Number |
        Sort-Object { [int]$_.Name } |
        ForEach-Object {
            $number = [int]$_.Name
            $trackName = ''
            if ($namesByNumber.ContainsKey($number)) { $trackName = @($namesByNumber[$number].Keys)[0] }
            [pscustomobject]@{
                Number = $number
                Name   = $trackName
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
