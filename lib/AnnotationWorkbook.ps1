<#
    Shared Excel-writing helpers for the AutoNyx annotation tools.
    Dot-source this file; it defines functions only and runs nothing.

    A "row" is any object with .Text, .Start, .End and (optionally) .Track.
#>

function New-ExcelApp {
    <# Starts a hidden Excel instance. Caller must pass it to Close-ExcelApp. #>
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    return $excel
}

function Close-ExcelApp {
    <#
    .PARAMETER KeepRunning
        Let go of the COM handle without quitting Excel - used after
        Show-Workbooks, so the window stays open for the user.
    #>
    param($Excel, [switch]$KeepRunning)
    if ($Excel) {
        if (-not $KeepRunning) { $Excel.Quit() }
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($Excel)
    }
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

function Show-Workbooks {
    <#
    .SYNOPSIS
        Makes the automation instance visible, opens the finished workbooks in
        it, and brings Excel to the front.
    .DESCRIPTION
        Deliberately not "Start-Process file.xlsx": the hotkey starts PowerShell
        with -WindowStyle Hidden, and a child launched from a hidden parent can
        inherit that hidden state, leaving the workbook open but invisible.
        Driving the instance we already own removes the guesswork.
    #>
    param(
        [Parameter(Mandatory)] $Excel,
        [Parameter(Mandatory)] [string[]]$Paths
    )

    # Hand the instance back to the user in a normal state - it was created
    # silent for automation.
    $Excel.DisplayAlerts = $true

    # Order matters. Setting Application.Visible first makes Excel show an empty
    # frame window, and the workbook then opens into a second window that stays
    # hidden - the file is "open" but nothing appears on screen. Opening while
    # the application is still invisible and revealing afterwards gives exactly
    # one real window per workbook.
    $firstHwnd = [IntPtr]::Zero
    foreach ($p in $Paths) {
        $book = $Excel.Workbooks.Open($p)
        $win  = $book.Windows.Item(1)
        $win.Visible = $true
        if ($firstHwnd -eq [IntPtr]::Zero) {
            try { $firstHwnd = [IntPtr]$win.Hwnd } catch { }
        }
    }

    $Excel.Visible = $true
    try { $Excel.WindowState = -4143 } catch { }   # xlNormal, in case it opened minimised

    if (-not ('AutoNyx.Win32' -as [type])) {
        Add-Type -Namespace AutoNyx -Name Win32 -MemberDefinition @'
[DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr hWnd);
[DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
'@
    }
    $hwnd = if ($firstHwnd -ne [IntPtr]::Zero) { $firstHwnd } else { [IntPtr]$Excel.Hwnd }
    [void][AutoNyx.Win32]::ShowWindow($hwnd, 9)          # SW_RESTORE
    [void][AutoNyx.Win32]::SetForegroundWindow($hwnd)
}

function Get-FreePath {
    <# Returns $Path, or "name (2).ext" etc. if it is taken. Never overwrites. #>
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $Path }

    $dir  = [System.IO.Path]::GetDirectoryName($Path)
    $base = [System.IO.Path]::GetFileNameWithoutExtension($Path)
    $ext  = [System.IO.Path]::GetExtension($Path)
    for ($i = 2; $i -lt 1000; $i++) {
        $candidate = Join-Path $dir "$base ($i)$ext"
        if (-not (Test-Path -LiteralPath $candidate)) { return $candidate }
    }
    throw "Could not find a free filename for $Path"
}

# The workbook layout, in column order. Every route writes all ten columns so
# that later pipeline steps see one consistent shape; columns with no Source
# are left empty for the annotator to fill in.
#
#   Source = the property read off each row, or $null for a column nobody fills yet
#   Format = Excel number format ('@' = text, so a leading "=" is not a formula)
#   Width  = fixed column width, or $null to size to contents
#
# Two layouts are kept side by side. They differ only in header names and column
# order - the same ten pieces of information are written either way - and each
# one matches a different filler, so an export can be opened in the filler
# without renaming anything:
#
#   SoupEE  headers and order of the "Soup EE" schema, read by CaptionFiller2
#           and CaptionFiller3 via lib\ExcelBridge2.ahk. Current default.
#   Legacy  the original layout, read by CaptionFiller.ahk via
#           lib\ExcelBridge.ahk. Kept so going back is one line, below.
$script:AnnotationLayouts = @{

    SoupEE = @(
        @{ Header = 'Caption Number'    ; Source = 'CaptionNumber' ; Format = '0'     ; Width = $null },
        @{ Header = 'TrackNumber'       ; Source = 'TrackNumber'   ; Format = '0'     ; Width = $null },
        @{ Header = 'Track'             ; Source = 'Track'         ; Format = '@'     ; Width = $null },
        @{ Header = 'TrackDescription'  ; Source = $null           ; Format = '@'     ; Width = 28 },
        @{ Header = 'StartTime(s)'      ; Source = 'Start'         ; Format = '0.000' ; Width = $null },
        @{ Header = 'EndTime(s)'        ; Source = 'End'           ; Format = '0.000' ; Width = $null },
        @{ Header = 'SourceVisibility'  ; Source = $null           ; Format = '@'     ; Width = 18 },
        @{ Header = 'SourceDescription' ; Source = $null           ; Format = '@'     ; Width = 28 },
        @{ Header = 'Prominence'        ; Source = $null           ; Format = '@'     ; Width = 14 },
        @{ Header = 'AnnotationText'    ; Source = 'Text'          ; Format = '@'     ; Width = 60 }
    )

    Legacy = @(
        @{ Header = 'Caption Number'     ; Source = 'CaptionNumber' ; Format = '0'     ; Width = $null },
        @{ Header = 'Track Number'       ; Source = 'TrackNumber'   ; Format = '0'     ; Width = $null },
        @{ Header = 'Track Name'         ; Source = 'Track'         ; Format = '@'     ; Width = $null },
        @{ Header = 'Track Description'  ; Source = $null           ; Format = '@'     ; Width = 28 },
        @{ Header = 'Caption'            ; Source = 'Text'          ; Format = '@'     ; Width = 60 },
        @{ Header = 'Start Time'         ; Source = 'Start'         ; Format = '0.000' ; Width = $null },
        @{ Header = 'End Time'           ; Source = 'End'           ; Format = '0.000' ; Width = $null },
        @{ Header = 'Source Visibility'  ; Source = $null           ; Format = '@'     ; Width = 18 },
        @{ Header = 'Source Description' ; Source = $null           ; Format = '@'     ; Width = 28 },
        @{ Header = 'Prominence'         ; Source = $null           ; Format = '@'     ; Width = 14 }
    )
}

# The layout both routes write. Change this single word to 'Legacy' to get the
# old headers back.
$script:AnnotationLayout = 'SoupEE'

function Write-AnnotationWorkbook {
    <#
    .SYNOPSIS
        Writes annotation rows to an .xlsx via Excel COM. Returns the path written.
    .DESCRIPTION
        Layout comes from $AnnotationLayouts above - edit that table to add,
        rename or reorder columns. A row property that is absent (e.g. Track on
        the .txt route, which has no track information) leaves its cell empty.
        Caption Number is generated here, numbering the rows as written.
    .PARAMETER Layout
        Which column layout to write. Defaults to $AnnotationLayout, so both
        routes stay in step unless a caller deliberately asks otherwise.
    #>
    param(
        [Parameter(Mandatory)] $Excel,
        [Parameter(Mandatory)] $Rows,
        [Parameter(Mandatory)] [string]$Path,
        [string]$SheetName = 'Annotations',
        [ValidateSet('SoupEE', 'Legacy')] [string]$Layout = $script:AnnotationLayout
    )

    $columns = $script:AnnotationLayouts[$Layout]
    $cols = $columns.Count

    # Copy into a plain array by hand: "@($Rows)" throws "Argument types do not
    # match" on a Generic.List[object] under Windows PowerShell 5.1.
    $items = New-Object System.Collections.ArrayList
    foreach ($r in $Rows) { [void]$items.Add($r) }
    $count = $items.Count

    $book  = $Excel.Workbooks.Add()
    $sheet = $book.Worksheets.Item(1)
    # Excel rejects these in sheet names, and truncates past 31 chars.
    $safe  = ($SheetName -replace '[\\/\?\*\[\]:]', '-')
    if ($safe.Length -gt 31) { $safe = $safe.Substring(0, 31) }
    if ($safe) { $sheet.Name = $safe }

    # Build a 2-D array and write it in one shot; cell-by-cell COM is glacial.
    $data = New-Object 'object[,]' ($count + 1), $cols
    for ($c = 0; $c -lt $cols; $c++) { $data[0, $c] = $columns[$c].Header }

    for ($i = 0; $i -lt $count; $i++) {
        $r = $items[$i]
        for ($c = 0; $c -lt $cols; $c++) {
            $source = $columns[$c].Source
            $value = switch ($source) {
                $null           { $null }                       # nobody fills this one yet
                'CaptionNumber' { $i + 1 }                      # generated, not read off the row
                default         { $r.PSObject.Properties[$source].Value }
            }
            # Parentheses required: "," binds tighter than "+" in PowerShell.
            $data[($i + 1), $c] = $value
        }
    }

    $all = $sheet.Range('A1').Resize($count + 1, $cols)
    for ($c = 0; $c -lt $cols; $c++) {
        $sheet.Range('A1').Offset(0, $c).Resize($count + 1, 1).NumberFormat = $columns[$c].Format
    }
    $all.Value2 = $data

    $header = $sheet.Range('A1').Resize(1, $cols)
    $header.Font.Bold = $true
    $header.Interior.Color = 15917529   # light grey

    # Cosmetic only, and FreezePanes needs the window to be active - never let
    # it cost us the conversion.
    try {
        $window = $book.Windows.Item(1)
        # [void] matters: Activate() emits an empty value that would otherwise
        # ride out of this function alongside the return value.
        [void]$window.Activate()
        $window.SplitRow = 1
        $window.SplitColumn = 0
        $window.FreezePanes = $true
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($window)
    } catch {
        Write-Verbose "Could not freeze the header row: $($_.Exception.Message)"
    }

    for ($c = 0; $c -lt $cols; $c++) {
        $column = $sheet.Columns.Item($c + 1)
        $column.WrapText = $false
        $width = $columns[$c].Width
        if ($width) {
            $column.ColumnWidth = $width
        } else {
            # Empty columns would shrink to nothing, so never go below the
            # header's own width.
            $column.AutoFit() | Out-Null
            $minimum = $columns[$c].Header.Length + 4
            if ($column.ColumnWidth -lt $minimum) { $column.ColumnWidth = $minimum }
        }
    }
    $all.AutoFilter() | Out-Null

    $outPath = Get-FreePath $Path
    $book.SaveAs($outPath, 51)   # 51 = xlOpenXMLWorkbook (.xlsx)
    $book.Close($false)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($sheet)
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($book)

    return $outPath
}
