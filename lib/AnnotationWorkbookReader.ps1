<#
    Reads annotation rows back out of an .xlsx, the other half of
    lib\AnnotationWorkbook.ps1 (which only writes them).

    The column layout is the "Soup EE" one - see $AnnotationLayouts there - and
    the sheets in play repeat all ten headers up to three times across the
    columns: the real data, a variant, then a block of TRUE/FALSE change flags.
    Mapping headers naively (last one wins) binds Track to the booleans, so the
    block-anchoring from lib\ExcelBridge2.ahk is repeated here.

    Dot-source this file; it defines functions only and runs nothing.
#>

# In sheet order. The first doubles as the anchor marking where a block begins.
$script:AnnotationSchema = @(
    'Caption Number', 'TrackNumber', 'Track', 'TrackDescription',
    'StartTime(s)', 'EndTime(s)', 'SourceVisibility',
    'SourceDescription', 'Prominence', 'AnnotationText'
)

# What a sheet must carry to count as annotations. TrackNumber is deliberately
# not here: workbooks vary, and a track named "1 Male Speech" carries its number
# in the name, which Get-AnnotationRows falls back to. Read-SoundNounFile is the
# one that ultimately needs a number, and Test-AnnotationRows says so by row.
$script:AnnotationRequired = @(
    'Caption Number', 'Track', 'StartTime(s)', 'EndTime(s)', 'AnnotationText'
)

# Property name written on each row object, per header.
$script:AnnotationFields = @{
    'Caption Number'    = 'CaptionNumber'
    'TrackNumber'       = 'TrackNumber'
    'Track'             = 'Track'
    'TrackDescription'  = 'TrackDescription'
    'StartTime(s)'      = 'Start'
    'EndTime(s)'        = 'End'
    'SourceVisibility'  = 'SourceVisibility'
    'SourceDescription' = 'SourceDescription'
    'Prominence'        = 'Prominence'
    'AnnotationText'    = 'Text'
}

# "1 Male Speech" or "1 - Male Speech". Written as escapes rather than literal
# dashes: Windows PowerShell 5.1 reads a BOM-less .ps1 as ANSI, which would
# corrupt an en-dash here and stop the pattern matching anything.
$script:TrackNamePrefix = '^\s*(?<num>[0-9]+)\s*[\u2013\u2014-]?\s+(?<name>.+)$'

function Find-AnnotationHeaderBlock {
    <#
    .SYNOPSIS
        Locates the schema on a worksheet. Returns @{ Columns; HeaderRow } or $null.
    .DESCRIPTION
        The block starts at the first occurrence of the anchor header and ends
        just before the anchor appears again; only that window is searched, and
        inside it the first occurrence of each header wins. Extra or interleaved
        columns within the block are tolerated.
    #>
    param([Parameter(Mandatory)]$Sheet)

    $used = $null
    try { $used = $Sheet.UsedRange } catch { return $null }
    if (-not $used) { return $null }

    $firstCol  = $used.Column
    $lastCol   = $used.Column + $used.Columns.Count - 1
    $headerRow = $used.Row
    if ($lastCol -lt $firstCol) { return $null }

    $values = $null
    try {
        $values = $Sheet.Range($Sheet.Cells($headerRow, $firstCol),
                               $Sheet.Cells($headerRow, $lastCol)).Value2
    } catch { return $null }
    if ($null -eq $values) { return $null }

    # One header cell comes back as a bare value rather than a 2-D array.
    $headers = @()
    if ($values -is [object[,]]) {
        for ($c = 1; $c -le $values.GetLength(1); $c++) {
            $headers += ([string]$values[1, $c]).Trim()
        }
    } else {
        $headers += ([string]$values).Trim()
    }

    $anchor = $script:AnnotationSchema[0]
    $start = 0
    $stop  = $headers.Count
    for ($i = 0; $i -lt $headers.Count; $i++) {
        if ($headers[$i] -ne $anchor) { continue }
        if ($start -eq 0) { $start = $i + 1 } else { $stop = $i; break }
    }
    if ($start -eq 0) { return $null }

    $columns = @{}
    for ($i = $start; $i -le $stop; $i++) {
        $header = $headers[$i - 1]
        if ($header -eq '' -or $columns.ContainsKey($header)) { continue }
        if ($script:AnnotationSchema -contains $header) {
            $columns[$header] = $i + $firstCol - 1
        }
    }

    foreach ($want in $script:AnnotationRequired) {
        if (-not $columns.ContainsKey($want)) { return $null }
    }

    return @{ Columns = $columns; HeaderRow = $headerRow; LastRow = $used.Row + $used.Rows.Count - 1 }
}

function ConvertTo-AnnotationDouble {
    <#
        Excel hands numeric cells back as doubles already; a cell formatted as
        text needs parsing, and invariantly, so a comma-decimal locale cannot
        read "1.400" as 1400.
    #>
    param($Value)
    if ($null -eq $Value) { return $null }
    if ($Value -is [double]) { return $Value }
    $text = ([string]$Value).Trim()
    if ($text -eq '') { return $null }
    $parsed = 0.0
    $styles = [System.Globalization.NumberStyles]::Float
    if ([double]::TryParse($text, $styles, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$parsed)) {
        return $parsed
    }
    return $null
}

function Get-AnnotationRows {
    <#
    .SYNOPSIS
        Rows from one bound sheet, in sheet order.
    .DESCRIPTION
        Each row carries CaptionNumber, TrackNumber, Track, TrackDescription,
        Start, End, SourceVisibility, SourceDescription, Prominence, Text and
        Row (the worksheet row, so problems can be reported where the user can
        find them).

        TrackNumber comes from its own column where the sheet has one. Where it
        does not, it is taken from the front of the Track name - the import
        names its tracks "1 Male Speech", so an export of imported tracks
        carries the number there - and the name keeps only what follows.

        A Track that already begins with this row's own number has that prefix
        removed, so a workbook exported from imported tracks can go round again
        without the number doubling into "1 1 Male Speech".
    #>
    param(
        [Parameter(Mandatory)]$Sheet,
        [Parameter(Mandatory)]$Block
    )

    $columns = $Block.Columns
    $first = $Block.HeaderRow + 1
    $last  = $Block.LastRow
    if ($last -lt $first) { return , @() }

    # One read for the whole block; cell-by-cell COM is glacial.
    $left  = ($columns.Values | Measure-Object -Minimum).Minimum
    $right = ($columns.Values | Measure-Object -Maximum).Maximum
    $grid  = $Sheet.Range($Sheet.Cells($first, $left), $Sheet.Cells($last, $right)).Value2

    $rows = New-Object System.Collections.Generic.List[object]
    $height = $last - $first + 1

    for ($r = 1; $r -le $height; $r++) {
        $record = @{ Row = $first + $r - 1 }
        foreach ($header in $columns.Keys) {
            $c = $columns[$header] - $left + 1
            $value = if ($grid -is [object[,]]) { $grid[$r, $c] } else { $grid }
            $record[$script:AnnotationFields[$header]] = $value
        }

        $start = ConvertTo-AnnotationDouble $record['Start']
        $end   = ConvertTo-AnnotationDouble $record['End']
        $track = ([string]$record['Track']).Trim()

        # A row with neither a time nor a track name is padding below the data.
        if ($null -eq $start -and $null -eq $end -and $track -eq '') { continue }

        $number = $null
        $numeric = ConvertTo-AnnotationDouble $record['TrackNumber']
        if ($null -ne $numeric) { $number = [int]$numeric }

        if ($null -eq $number) {
            if ($track -match $script:TrackNamePrefix) {
                $number = [int]$Matches['num']
                $track  = $Matches['name'].Trim()
            }
        }
        elseif ($track -match $script:TrackNamePrefix -and [int]$Matches['num'] -eq $number) {
            $track = $Matches['name'].Trim()
        }

        $caption = ConvertTo-AnnotationDouble $record['CaptionNumber']

        $rows.Add([pscustomobject]@{
            CaptionNumber     = if ($null -ne $caption) { [int]$caption } else { $rows.Count + 1 }
            TrackNumber       = $number
            Track             = $track
            TrackDescription  = ([string]$record['TrackDescription']).Trim()
            Start             = $start
            End               = $end
            SourceVisibility  = ([string]$record['SourceVisibility']).Trim()
            SourceDescription = ([string]$record['SourceDescription']).Trim()
            Prominence        = ([string]$record['Prominence']).Trim()
            Text              = ([string]$record['Text']).Trim()
            Row               = $record.Row
        })
    }

    # ToArray rather than @($rows): Windows PowerShell 5.1 throws "Argument
    # types do not match" wrapping a Generic.List[object] that way.
    return , $rows.ToArray()
}

function Read-AnnotationWorkbook {
    <#
    .SYNOPSIS
        Opens an .xlsx, finds the annotation sheet, and returns its rows.
    .DESCRIPTION
        Returns [pscustomobject] @{ Path; Sheet; Rows }.

        A workbook usually holds several sheets carrying this schema -
        "Completion" and "Refinement" - alongside sheets that do not, such as
        "Summary". Exactly one match binds silently; several ask, through
        ChooseSheet, while the workbook is still open, so Excel is started once.

        A sheet is only a match if it has captions under the headers. A prepared
        but still empty pass is not a choice worth asking about, and the sheet
        picked would have thrown "holds no captions" a moment later anyway.
    .PARAMETER Sheet
        Bind this sheet instead of searching. Compared case-sensitively, since
        these names sometimes differ only by a trailing space.
    .PARAMETER ChooseSheet
        Called when more than one sheet matches, with one object per candidate
        carrying Name and Count (its captions). Must return one of the names, or
        $null to cancel.
    #>
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)][string]$Path,
        [string]$Sheet,
        [scriptblock]$ChooseSheet
    )

    if (-not (Test-Path -LiteralPath $Path)) { throw "Cannot find the workbook:`n$Path" }
    $full = (Resolve-Path -LiteralPath $Path).ProviderPath

    $excel = $null
    $book  = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        # Read-only, and no links refreshed: this only reads, and a workbook that
        # stops to ask about anything would hang a hidden process.
        $book = $excel.Workbooks.Open($full, 0, $true)

        # Every candidate is read here, before anything is chosen: an empty one
        # can then be left out of the question entirely, the picker can say how
        # many captions each holds, and the sheet that wins does not have to be
        # read a second time. One block read per sheet, so it is cheap.
        #
        # Not $matches: that is the automatic variable -match writes into, and
        # PowerShell variable names are case-insensitive, so the name is taken.
        $found = New-Object System.Collections.Specialized.OrderedDictionary
        $blank = New-Object System.Collections.Generic.List[string]
        foreach ($ws in $book.Worksheets) {
            $block = Find-AnnotationHeaderBlock -Sheet $ws
            if (-not $block) { continue }
            $rows = Get-AnnotationRows -Sheet $ws -Block $block
            if (@($rows).Count -eq 0) { $blank.Add($ws.Name); continue }
            $found[$ws.Name] = $rows
        }

        if ($found.Count -eq 0) {
            if ($blank.Count -gt 0) {
                throw ("$([System.IO.Path]::GetFileName($full)) has no sheet with captions on it." +
                       "`n`nThese carry the annotation columns but nothing underneath them: " +
                       ($blank -join ', ') + '.')
            }
            throw ("$([System.IO.Path]::GetFileName($full)) has no sheet carrying the annotation " +
                   "columns.`n`nA sheet needs at least: " +
                   ($script:AnnotationRequired -join ', ') + '.')
        }

        $names = @($found.Keys)
        $name = $null
        if ($Sheet) {
            foreach ($key in $names) { if ($key -ceq $Sheet) { $name = $key; break } }
            if (-not $name) {
                # Named an empty one: say so, rather than listing the sheets and
                # leaving the user to wonder where theirs went.
                foreach ($key in $blank) {
                    if ($key -ceq $Sheet) {
                        throw ("$([System.IO.Path]::GetFileName($full)) sheet '$Sheet' carries the " +
                               "annotation columns but no captions underneath them.")
                    }
                }
                throw ("$([System.IO.Path]::GetFileName($full)) has no annotation sheet called " +
                       "'$Sheet'.`n`nIt has: " + ($names -join ', '))
            }
        }
        elseif ($names.Count -eq 1) {
            $name = $names[0]
        }
        elseif ($ChooseSheet) {
            $choices = @(foreach ($key in $names) {
                [pscustomobject]@{ Name = $key; Count = @($found[$key]).Count }
            })
            $name = & $ChooseSheet $choices
            if (-not $name) { return $null }
            if (-not $found.Contains($name)) { throw "No annotation sheet called '$name'." }
        }
        else {
            throw ("$([System.IO.Path]::GetFileName($full)) has $($names.Count) annotation " +
                   "sheets - " + ($names -join ', ') +
                   ".`n`nSay which with -Sheet.")
        }

        return [pscustomobject]@{
            Path  = $full
            Sheet = $name
            Rows  = $found[$name]
        }
    }
    finally {
        if ($book)  { $book.Close($false); [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book) }
        if ($excel) { $excel.Quit();       [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel) }
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

function Test-AnnotationRows {
    <#
    .SYNOPSIS
        Everything that can be checked before spending a model call. Returns the
        problems as strings; empty means the rows are fit to build a file from.
    .DESCRIPTION
        The one that matters is the last: one track number must carry one track
        name. A workbook that breaks it cannot become label tracks at all, since
        there would be nothing to call the track - and it is exactly what the
        import rejected the hand-made files for, so catching it here means the
        user hears about it before waiting on Claude rather than after.
    #>
    param([Parameter(Mandatory)]$Rows)

    $problems = New-Object System.Collections.Generic.List[string]

    foreach ($row in $Rows) {
        $where = "row $($row.Row)"
        if ($null -eq $row.TrackNumber) {
            $problems.Add("  ${where}: no track number, and '$($row.Track)' does not start with one")
        }
        if ($row.Track -eq '') { $problems.Add("  ${where}: no track name") }
        if ($row.Track -match "`t") { $problems.Add("  ${where}: the track name contains a tab") }
        if ($null -eq $row.Start -or $null -eq $row.End) {
            $problems.Add("  ${where}: start or end time is missing or not a number")
        }
        elseif ($row.End -lt $row.Start) {
            $problems.Add(("  {0}: ends before it starts ({1:0.000} to {2:0.000})" -f $where, $row.Start, $row.End))
        }
    }

    # Ordinal, so that "Male speech" and "Male Speech" count as two names: which
    # casing is the real one is not something to decide here.
    $namesByNumber = @{}
    foreach ($row in $Rows) {
        if ($null -eq $row.TrackNumber -or $row.Track -eq '') { continue }
        if (-not $namesByNumber.ContainsKey($row.TrackNumber)) {
            $namesByNumber[$row.TrackNumber] =
                New-Object System.Collections.Specialized.OrderedDictionary ([System.StringComparer]::Ordinal)
        }
        $seen = $namesByNumber[$row.TrackNumber]
        if (-not $seen.Contains($row.Track)) { $seen[$row.Track] = $row.Row }
    }
    foreach ($number in ($namesByNumber.Keys | Sort-Object)) {
        $seen = $namesByNumber[$number]
        if ($seen.Count -le 1) { continue }
        $detail = @(foreach ($key in $seen.Keys) { "row $($seen[$key]) '$key'" }) -join ', '
        $problems.Add("  track ${number} has more than one name: $detail")
    }

    return , $problems.ToArray()
}
