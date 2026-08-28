Attribute VB_Name = "modFormatWorkbook"
Option Explicit

' ===========================================================================
'  CC - applies the house formatting to every worksheet in the workbook.
'
'  Per sheet:
'    1. Consolas 8 pt, top-aligned, across the whole sheet
'    2. where the sheet has column headers (row 1 is not blank):
'         - row 1 bolded and highlighted yellow, across the used columns
'         - AutoFilter applied over the header row
'         - row 1 frozen
'
'  Sheets with no headers are still formatted, but get no highlight, filter
'  or freeze. Protected sheets are skipped and reported at the end.
'
'  Re-runnable: reapplying is a no-op on an already-formatted workbook.
' ===========================================================================

' The row treated as the header row.
Private Const HEADER_ROW     As Long = 1

' Body formatting.
Private Const BODY_FONT      As String = "Consolas"
Private Const BODY_FONT_SIZE As Long = 8


Public Sub CC()
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim entrySheet As Object
    Dim skipped As String

    If ActiveWorkbook Is Nothing Then Exit Sub
    Set wb = ActiveWorkbook
    Set entrySheet = ActiveSheet

    On Error GoTo Cleanup
    Application.ScreenUpdating = False
    wb.Activate

    For Each ws In wb.Worksheets
        If ws.ProtectContents Then
            skipped = skipped & vbLf & "  " & ws.Name
        Else
            FormatSheet ws
        End If
    Next ws

Cleanup:
    RestoreActiveSheet entrySheet
    Application.ScreenUpdating = True

    If Err.Number <> 0 Then
        MsgBox "CC failed: " & Err.Description, vbExclamation
        Err.Clear
    ElseIf Len(skipped) > 0 Then
        MsgBox "Skipped these protected sheets:" & vbLf & skipped, vbInformation
    End If
End Sub


' --- per sheet -------------------------------------------------------------

Private Sub FormatSheet(ws As Worksheet)
    Dim lastRow As Long, lastCol As Long

    ApplyBodyFormat ws
    If Not HasColumnHeaders(ws) Then Exit Sub

    lastRow = LastUsedRow(ws)
    lastCol = LastUsedColumn(ws)

    HighlightHeaderRow ws, lastCol
    ApplyHeaderAutoFilter ws, lastRow, lastCol
    FreezeHeaderRow ws
End Sub

' Consolas 8 pt, top-aligned, no wrapping. Horizontal alignment is reset to
' General rather than forced left, so text still sits left and numbers still
' sit right. Merged cells are left alone - this runs over the whole workbook,
' and unmerging would wreck any sheet that is laid out rather than tabular.
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

' Only the used columns - filling the whole row out to XFD looks wrong the
' moment anyone scrolls right.
Private Sub HighlightHeaderRow(ws As Worksheet, ByVal lastCol As Long)
    With ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(HEADER_ROW, lastCol))
        .Interior.Color = vbYellow
        .Font.Bold = True
    End With
End Sub

Private Sub ApplyHeaderAutoFilter(ws As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long)
    ' A ListObject carries its own filter, and a sheet-level AutoFilter over
    ' the same cells is refused - just make sure the table's is showing.
    If ws.ListObjects.Count > 0 Then
        ShowTableFilters ws
        Exit Sub
    End If

    ' A one-row range makes Excel guess the extent, so always span two rows.
    If lastRow < HEADER_ROW + 1 Then lastRow = HEADER_ROW + 1

    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range(ws.Cells(HEADER_ROW, 1), ws.Cells(lastRow, lastCol)).AutoFilter
End Sub

Private Sub ShowTableFilters(ws As Worksheet)
    Dim tbl As ListObject

    For Each tbl In ws.ListObjects
        If tbl.ShowHeaders Then tbl.ShowAutoFilter = True
    Next tbl
End Sub

' Freezing is a window operation, so the sheet has to be the active one -
' which rules out hidden sheets. Scroll home first: FreezePanes splits
' relative to the top-left visible cell, not to A1.
Private Sub FreezeHeaderRow(ws As Worksheet)
    If ws.Visible <> xlSheetVisible Then Exit Sub

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

' Headers are present when row 1 holds anything at all.
Private Function HasColumnHeaders(ws As Worksheet) As Boolean
    HasColumnHeaders = (Application.CountA(ws.Rows(HEADER_ROW)) > 0)
End Function

' Puts the caller back where they started, without upsetting the run if that
' sheet has since been hidden.
Private Sub RestoreActiveSheet(entrySheet As Object)
    On Error Resume Next
    If Not entrySheet Is Nothing Then entrySheet.Activate
    On Error GoTo 0
End Sub

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
