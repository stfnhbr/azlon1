Attribute VB_Name = "modCaptions"
Option Explicit

' ===========================================================================
'  HYX - tidies a caption export and prefixes it with a "Caption Number"
'        column that numbers the existing caption rows, and only those.
'
'  What it does, in order:
'    1. normalises sheet-wide formatting (font, alignment, no merged cells)
'    2. inserts column A, "Caption Number", unless the sheet already has one
'    3. numbers every row that carries data; blank rows stay blank, and
'       nothing below the last used row is touched
'    4. centres Caption Number and Track Number, and formats Start/End time
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

    On Error GoTo Cleanup
    Application.ScreenUpdating = False

    NormaliseSheet ws
    EnsureCaptionColumn ws

    lastCol = LastUsedColumn(ws)
    WriteCaptionNumbers ws, lastRow, lastCol
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

' Strips the export's own formatting back to a single readable baseline.
Private Sub NormaliseSheet(ws As Worksheet)
    With ws.Cells
        .MergeCells = False
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

Private Sub FormatColumns(ws As Worksheet)
    ' Caption Number and Track Number, centred.
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
