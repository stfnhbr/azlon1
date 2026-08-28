Attribute VB_Name = "modCaptions"
Option Explicit

' ===========================================================================
'  HYX - tidies a caption export and prefixes it with a "Caption Number"
'        column that numbers the existing caption rows, and only those.
'
'  What it does, in order:
'    1. sets the sheet to Consolas 8 pt, top-aligned, and unmerges it - the
'       active sheet only, never the rest of the workbook
'    2. inserts column A, "Caption Number", unless the sheet already has one
'    3. numbers every row that carries data; blank rows stay blank, and
'       nothing below the last used row is touched
'    4. centres Caption Number and Track Number, converts Start/End time from
'       text to real numbers and formats them to three decimals
'    5. bolds and highlights the header row, reapplies the AutoFilter across
'       the full table, and freezes the header row
'
'  Re-runnable: a second run renumbers in place instead of inserting a
'  second column.
' ===========================================================================

' Header row and the header text that marks an already-numbered sheet.
Private Const HEADER_ROW     As Long = 1
Private Const CAPTION_HEADER As String = "Caption Number"

' Body formatting.
Private Const BODY_FONT      As String = "Consolas"
Private Const BODY_FONT_SIZE As Long = 8
Private Const TIME_FORMAT    As String = "0.000"

' Column positions in the raw export, before the Caption Number column is
' inserted. Everything downstream goes through SheetColumn(), so these are
' the only numbers to change if the export layout moves.
Private Const RAW_TRACK_COL  As Long = 1   ' Track Number
Private Const RAW_TIME_FIRST As Long = 4   ' Start time
Private Const RAW_TIME_LAST  As Long = 5   ' End time


Public Sub HYX()
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long

    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    Set ws = ActiveSheet

    ' Measured across the whole sheet rather than off one column: the
    ' optional columns (Track Description, Prominence, ...) are empty until
    ' the annotator fills them in, so any single column can come up short.
    lastRow = LastUsedRow(ws)
    If lastRow < HEADER_ROW Then Exit Sub       ' empty sheet, nothing to number

    If ws.ProtectContents Then
        MsgBox "'" & ws.Name & "' is protected - unprotect it and run HYX again.", _
               vbExclamation
        Exit Sub
    End If

    On Error GoTo Cleanup
    Application.ScreenUpdating = False

    ApplyBodyFormat ws
    UnmergeSheet ws
    EnsureCaptionColumn ws

    lastCol = LastUsedColumn(ws)
    WriteCaptionNumbers ws, lastRow, lastCol
    ConvertTimesToNumbers ws, lastRow
    FormatColumns ws
    HighlightHeaderRow ws, lastCol
    ApplyAutoFilter ws, lastRow, lastCol
    FreezeHeaderRow ws

    ws.Range("A1").Select

Cleanup:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "HYX failed: " & Err.Description, vbExclamation
        Err.Clear
    End If
End Sub


' --- steps -----------------------------------------------------------------

' Consolas 8 pt, top-aligned, on the export sheet only. HYX never touches
' another sheet: resetting horizontal alignment workbook-wide would undo the
' centring a previous run put on another sheet's columns A and B, and those
' hold numbers, so General alignment shows them right-aligned.
'
' Horizontal alignment is reset to General rather than forced left, so text
' still sits left and numbers still sit right.
Private Sub ApplyBodyFormat(ws As Worksheet)
    With ws.Cells
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .IndentLevel = 0
        .ShrinkToFit = False
        With .Font
            .Name = BODY_FONT
            .Size = BODY_FONT_SIZE
            .Underline = xlUnderlineStyleNone
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
        End With
    End With
End Sub

Private Sub UnmergeSheet(ws As Worksheet)
    ws.Cells.MergeCells = False
End Sub

' Inserts the Caption Number column, unless a previous run already did.
Private Sub EnsureCaptionColumn(ws As Worksheet)
    If StrComp(Trim$(CStr(ws.Cells(HEADER_ROW, 1).Value)), _
               CAPTION_HEADER, vbTextCompare) = 0 Then Exit Sub

    ws.Columns(1).Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    ws.Cells(HEADER_ROW, 1).Value = CAPTION_HEADER
End Sub

' Numbers the rows that carry data. The whole block is read in one go and
' written back in one go - no AutoFill, so there is no seed range to size
' the fill from and no stale destination to inherit, and a sheet with a
' single data row still numbers correctly.
Private Sub WriteCaptionNumbers(ws As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long)
    Dim body As Variant, numbers() As Variant
    Dim firstRow As Long, r As Long, n As Long

    firstRow = HEADER_ROW + 1
    If lastRow < firstRow Then Exit Sub

    ' Column A is the numbers themselves, so the data starts at B.
    If lastCol < 2 Then lastCol = 2
    body = ReadBlock(ws.Range(ws.Cells(firstRow, 2), ws.Cells(lastRow, lastCol)))

    ReDim numbers(1 To UBound(body, 1), 1 To 1)
    For r = 1 To UBound(body, 1)
        If RowHasData(body, r) Then
            n = n + 1
            numbers(r, 1) = n
        End If
        ' Blank rows keep an Empty element, which writes back as a blank cell.
    Next r

    ws.Cells(firstRow, 1).Resize(UBound(numbers, 1), 1).Value = numbers
End Sub

' A number format does nothing to a cell holding text, and the export writes
' Start/End as text often enough that "0.000" silently had no effect. Coerce
' anything that reads as a number first; genuine numbers are left as they are,
' and columns carrying formulas are left alone entirely.
Private Sub ConvertTimesToNumbers(ws As Worksheet, ByVal lastRow As Long)
    Dim timeCells As Range
    Dim block As Variant
    Dim r As Long, c As Long
    Dim parsedValue As Double
    Dim parsed As Boolean, changed As Boolean

    If lastRow < HEADER_ROW + 1 Then Exit Sub

    Set timeCells = ws.Range(ws.Cells(HEADER_ROW + 1, SheetColumn(RAW_TIME_FIRST)), _
                             ws.Cells(lastRow, SheetColumn(RAW_TIME_LAST)))
    ' HasFormula is Null when the range mixes formulas and values.
    If IsNull(timeCells.HasFormula) Then Exit Sub
    If timeCells.HasFormula Then Exit Sub

    block = ReadBlock(timeCells)
    For r = 1 To UBound(block, 1)
        For c = 1 To UBound(block, 2)
            If VarType(block(r, c)) = vbString Then
                parsedValue = AsNumber(block(r, c), parsed)
                If parsed Then
                    block(r, c) = parsedValue
                    changed = True
                End If
            End If
        Next c
    Next r

    If changed Then timeCells.Value = block
End Sub

Private Sub FormatColumns(ws As Worksheet)
    ' Column A (Caption Number) and column B (Track Number), centred.
    With ColumnSpan(ws, 1, SheetColumn(RAW_TRACK_COL))
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
    End With

    ' Start / End times, three decimals.
    ColumnSpan(ws, SheetColumn(RAW_TIME_FIRST), _
                   SheetColumn(RAW_TIME_LAST)).NumberFormat = TIME_FORMAT

    ws.Columns(1).AutoFit
End Sub

' The whole header row, not just the Caption Number cell - and only the used
' columns, since filling out to XFD looks wrong the moment anyone scrolls
' right.
Private Sub HighlightHeaderRow(ws As Worksheet, ByVal lastCol As Long)
    With ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(HEADER_ROW, lastCol))
        .Font.Bold = True
        .Interior.Color = vbYellow
    End With
End Sub

' The export ships its own AutoFilter over its own columns, and the insert
' pushed that range one column right - drop it and reapply across everything
' so Caption Number gets a dropdown too.
Private Sub ApplyAutoFilter(ws As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long)
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lastRow, lastCol)).AutoFilter
End Sub


' Freezing is a window operation, so it acts on the active sheet - which ws
' is, by construction. Scroll home first: FreezePanes splits relative to the
' top-left visible cell, not to A1.
Private Sub FreezeHeaderRow(ws As Worksheet)
    ws.Activate
    If ActiveWindow Is Nothing Then Exit Sub

    With ActiveWindow
        .FreezePanes = False
        .SplitRow = 0
        .SplitColumn = 0
        .ScrollRow = 1
        .ScrollColumn = 1
        .SplitRow = HEADER_ROW
        .FreezePanes = True
    End With
End Sub


' --- helpers ---------------------------------------------------------------

' Raw export column -> its position once Caption Number sits in front of it.
Private Function SheetColumn(ByVal rawColumn As Long) As Long
    SheetColumn = rawColumn + 1
End Function

' The whole of columns first..last.
Private Function ColumnSpan(ws As Worksheet, ByVal firstColumn As Long, _
                            ByVal lastColumn As Long) As Range
    Set ColumnSpan = ws.Range(ws.Cells(1, firstColumn), _
                              ws.Cells(1, lastColumn)).EntireColumn
End Function

' Last used row / column of the sheet, 0 when the sheet is empty.
Private Function LastUsedRow(ws As Worksheet) As Long
    Dim found As Range
    Set found = FindLast(ws, xlByRows)
    If Not found Is Nothing Then LastUsedRow = found.Row
End Function

Private Function LastUsedColumn(ws As Worksheet) As Long
    Dim found As Range
    Set found = FindLast(ws, xlByColumns)
    If Not found Is Nothing Then LastUsedColumn = found.Column
End Function

' Searches backwards from A1 to land on the final used cell in that order.
Private Function FindLast(ws As Worksheet, ByVal order As XlSearchOrder) As Range
    Set FindLast = ws.Cells.Find(What:="*", After:=ws.Range("A1"), _
                                 LookIn:=xlFormulas, LookAt:=xlPart, _
                                 SearchOrder:=order, SearchDirection:=xlPrevious)
End Function

' Text that reads as a number, with ok set when it did. The separator is
' localised first, so an export written on a machine with different regional
' settings still parses - "12,345" and "12.345" both come back as 12.345.
Private Function AsNumber(ByVal cellValue As Variant, ByRef ok As Boolean) As Double
    Dim raw As String

    ok = False
    raw = LocaliseDecimal(Trim$(CStr(cellValue)))
    If Len(raw) = 0 Then Exit Function

    If IsNumeric(raw) Then
        AsNumber = CDbl(raw)
        ok = True
    End If
End Function

' Swaps whichever decimal separator the text carries for the local one. Only
' applied when there is exactly one separator, so a thousands separator is
' never mistaken for a decimal point.
Private Function LocaliseDecimal(ByVal raw As String) As String
    Dim separators As Long

    separators = Len(raw) - Len(Replace(Replace(raw, ".", ""), ",", ""))
    If separators <> 1 Then
        LocaliseDecimal = raw
    Else
        LocaliseDecimal = Replace(Replace(raw, ".", DecimalSeparator), _
                                  ",", DecimalSeparator)
    End If
End Function

Private Function DecimalSeparator() As String
    DecimalSeparator = Application.International(xlDecimalSeparator)
End Function

' Range values as a 2-D array, including the single-cell case where
' Range.Value would hand back a bare scalar.
Private Function ReadBlock(rng As Range) As Variant
    Dim block As Variant

    If rng.Cells.CountLarge = 1 Then
        ReDim block(1 To 1, 1 To 1)
        block(1, 1) = rng.Value
    Else
        block = rng.Value
    End If

    ReadBlock = block
End Function

' A row counts as data if any cell holds something other than blank or "".
Private Function RowHasData(body As Variant, ByVal r As Long) As Boolean
    Dim c As Long, cell As Variant

    For c = 1 To UBound(body, 2)
        cell = body(r, c)
        If IsError(cell) Then
            RowHasData = True
            Exit Function
        ElseIf Not IsEmpty(cell) Then
            If Len(Trim$(CStr(cell))) > 0 Then
                RowHasData = True
                Exit Function
            End If
        End If
    Next c
End Function
