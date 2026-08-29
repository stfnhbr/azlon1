Attribute VB_Name = "modFormatWorkbook"
Option Explicit

' ===========================================================================
'  CC - applies the house formatting to the active worksheet.
'
'  What it does:
'    1. Consolas 8 pt, top-aligned, across the whole sheet
'    2. where the sheet has column headers (row 1 is not blank):
'         - row 1 bolded and highlighted yellow, across the used columns
'         - AutoFilter applied over the header row
'         - row 1 frozen
'
'  A sheet with no headers is still formatted, but gets no highlight, filter
'  or freeze. CC never touches another sheet: resetting alignment across the
'  workbook would undo the centring another sheet's columns rely on.
'
'  Re-runnable: reapplying is a no-op on an already-formatted sheet.
' ===========================================================================

' The row treated as the header row.
Private Const HEADER_ROW     As Long = 1

' Body formatting.
Private Const BODY_FONT      As String = "Consolas"
Private Const BODY_FONT_SIZE As Long = 8


Public Sub CC()
    Dim ws As Worksheet

    If TypeName(ActiveSheet) <> "Worksheet" Then Exit Sub
    Set ws = ActiveSheet

    If ws.ProtectContents Then
        MsgBox "'" & ws.Name & "' is protected - unprotect it and run CC again.", _
               vbExclamation
        Exit Sub
    End If

    On Error GoTo Cleanup
    Application.ScreenUpdating = False

    FormatSheet ws

Cleanup:
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "CC failed: " & Err.Description, vbExclamation
        Err.Clear
    End If
End Sub


' --- steps -----------------------------------------------------------------

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
' sit right. Merged cells are left alone - CC is a formatting pass, and
' unmerging would wreck a sheet that is laid out rather than tabular.
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

' Headers are present when row 1 holds anything at all.
Private Function HasColumnHeaders(ws As Worksheet) As Boolean
    HasColumnHeaders = (Application.CountA(ws.Rows(HEADER_ROW)) > 0)
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
