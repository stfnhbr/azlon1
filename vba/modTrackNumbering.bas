Attribute VB_Name = "modTrackNumbering"
Option Explicit

'==============================================================================
' modTrackNumbering
'
' Renumbers the TrackNumber column of a captioning sheet so that tracks are
' numbered 1..N in the chronological order of each track's FIRST caption.
'
' A "track" is a group of caption rows that belong to the same sound source.
' Each track's sort key is the smallest StartTime(s) found on any of its rows;
' ties are broken by the first row on which the track appears.
'
' Public entry points
'   RenumberTracksByFirstCaption ....... renumber the sheet named SHEET_NAME
'   RenumberTracksOnActiveSheet ........ renumber whichever sheet is active
'   UndoLastTrackRenumber .............. restore the values from the last run
'
' Notes
'   * Running a macro clears Excel's undo stack, so Ctrl+Z will NOT undo this.
'     The original numbers are saved to a very-hidden backup sheet instead;
'     use UndoLastTrackRenumber to restore them.
'   * Column positions are found by header text, not hard-coded letters, so
'     inserting or reordering columns is safe. Header matching ignores case,
'     spaces, underscores and punctuation ("StartTime(s)", "Start Time (s)"
'     and "start_time" all match).
'   * No external references are required (no Scripting.Dictionary), so this
'     runs on Excel for Windows and Excel for Mac.
'==============================================================================

'--- Settings -----------------------------------------------------------------

' Worksheet to process with RenumberTracksByFirstCaption.
Private Const SHEET_NAME As String = "Refinement"

' Row holding the column headers. Data is assumed to start on the next row.
Private Const HEADER_ROW As Long = 1

' How caption rows are grouped into tracks:
'   "TrackNumber" - group by the existing TrackNumber, falling back to the
'                   Track name when TrackNumber is blank (default).
'   "Track"       - group by the Track name only. Use this when the existing
'                   numbers are unreliable or missing.
Private Const GROUP_BY As String = "TrackNumber"

' Also resequence the "Caption Number" column to 1..N in row order.
Private Const RESEQUENCE_CAPTION_NUMBERS As Boolean = False

' Show the old -> new mapping when the macro finishes.
Private Const SHOW_SUMMARY As Boolean = True

' Name of the very-hidden sheet used to store undo data.
Private Const BACKUP_SHEET As String = "_TrackNumberBackup"

'--- Public entry points ------------------------------------------------------

Public Sub RenumberTracksByFirstCaption()
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(SHEET_NAME)
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "No worksheet named """ & SHEET_NAME & """ was found in " & _
               ActiveWorkbook.Name & "." & vbCrLf & vbCrLf & _
               "Either rename the sheet, change the SHEET_NAME constant at the " & _
               "top of this module, or run RenumberTracksOnActiveSheet instead.", _
               vbExclamation, "Renumber tracks"
        Exit Sub
    End If

    RenumberTracks ws
End Sub

Public Sub RenumberTracksOnActiveSheet()
    If TypeOf ActiveSheet Is Worksheet Then
        RenumberTracks ActiveSheet
    Else
        MsgBox "Select a worksheet first.", vbExclamation, "Renumber tracks"
    End If
End Sub

'--- Core routine -------------------------------------------------------------

Private Sub RenumberTracks(ws As Worksheet)
    Dim colTrackNum As Long, colTrack As Long, colStart As Long, colCaption As Long
    Dim firstRow As Long, lastRow As Long, lastCol As Long
    Dim data As Variant, outNums As Variant, outCaptions As Variant
    Dim keys() As String, labels() As String
    Dim minStart() As Double, hasStart() As Boolean, firstSeen() As Long
    Dim oldLabel() As String
    Dim order() As Long
    Dim trackCount As Long, idx As Long
    Dim r As Long, i As Long
    Dim key As String, label As String
    Dim startVal As Variant
    Dim changed As Long, skipped As Long, captionNo As Long
    Dim prevScreen As Boolean, prevEvents As Boolean, prevCalc As XlCalculation
    Dim stateSaved As Boolean

    On Error GoTo Fail

    If ws.ProtectContents Then
        MsgBox "Sheet """ & ws.Name & """ is protected. Unprotect it and run the " & _
               "macro again.", vbExclamation, "Renumber tracks"
        Exit Sub
    End If

    '--- Locate the columns we need by header text --------------------------
    lastCol = HeaderLastColumn(ws)

    colTrackNum = FindColumn(ws, lastCol, "tracknumber|tracknum|trackno|trackid")
    colTrack = FindColumn(ws, lastCol, "track|trackname|tracklabel")
    colStart = FindColumn(ws, lastCol, "starttimes|starttime|starttimesec|starttimeseconds|start|startsec|startseconds")
    colCaption = FindColumn(ws, lastCol, "captionnumber|captionno|captionnum|caption")

    If colTrackNum = 0 Then
        MsgBox "No ""TrackNumber"" column found in row " & HEADER_ROW & " of """ & _
               ws.Name & """.", vbExclamation, "Renumber tracks"
        Exit Sub
    End If
    If colStart = 0 Then
        MsgBox "No ""StartTime(s)"" column found in row " & HEADER_ROW & " of """ & _
               ws.Name & """.", vbExclamation, "Renumber tracks"
        Exit Sub
    End If
    If LCase$(GROUP_BY) = "track" And colTrack = 0 Then
        MsgBox "GROUP_BY is set to ""Track"" but no ""Track"" column was found in " & _
               "row " & HEADER_ROW & " of """ & ws.Name & """.", _
               vbExclamation, "Renumber tracks"
        Exit Sub
    End If

    '--- Work out the extent of the data ------------------------------------
    firstRow = HEADER_ROW + 1
    lastRow = LastDataRow(ws, colTrackNum, colTrack, colStart)

    If lastRow < firstRow Then
        MsgBox "No caption rows found below row " & HEADER_ROW & " on """ & _
               ws.Name & """.", vbInformation, "Renumber tracks"
        Exit Sub
    End If

    data = ws.Range(ws.Cells(firstRow, 1), ws.Cells(lastRow, lastCol)).Value

    ReDim keys(1 To lastRow - firstRow + 1)
    ReDim labels(1 To lastRow - firstRow + 1)
    ReDim oldLabel(1 To lastRow - firstRow + 1)
    ReDim minStart(1 To lastRow - firstRow + 1)
    ReDim hasStart(1 To lastRow - firstRow + 1)
    ReDim firstSeen(1 To lastRow - firstRow + 1)
    trackCount = 0

    '--- Pass 1: collect the tracks and each track's earliest start time -----
    For r = 1 To UBound(data, 1)
        key = TrackKey(data, r, colTrackNum, colTrack)

        If Len(key) = 0 Then
            skipped = skipped + 1
        Else
            idx = IndexOfKey(keys, trackCount, key)
            If idx = 0 Then
                trackCount = trackCount + 1
                idx = trackCount
                keys(idx) = key
                labels(idx) = TrackLabel(data, r, colTrackNum, colTrack)
                oldLabel(idx) = CellText(data, r, colTrackNum)
                firstSeen(idx) = r
                hasStart(idx) = False
                minStart(idx) = 0
            End If

            startVal = CellValue(data, r, colStart)
            If IsNumericValue(startVal) Then
                If Not hasStart(idx) Then
                    hasStart(idx) = True
                    minStart(idx) = CDbl(startVal)
                ElseIf CDbl(startVal) < minStart(idx) Then
                    minStart(idx) = CDbl(startVal)
                End If
            End If
        End If
    Next r

    If trackCount = 0 Then
        MsgBox "No caption rows with a track could be identified on """ & _
               ws.Name & """.", vbInformation, "Renumber tracks"
        Exit Sub
    End If

    '--- Pass 2: order tracks by first caption, then assign 1..N ------------
    order = SortedTrackOrder(trackCount, minStart, hasStart, firstSeen)

    ' newNumber(idx) is the number assigned to the track stored at keys(idx)
    Dim newNumber() As Long
    ReDim newNumber(1 To trackCount)
    For i = 1 To trackCount
        newNumber(order(i)) = i
    Next i

    '--- Pass 3: build the replacement column(s) ----------------------------
    ReDim outNums(1 To UBound(data, 1), 1 To 1)
    If RESEQUENCE_CAPTION_NUMBERS And colCaption > 0 Then
        ReDim outCaptions(1 To UBound(data, 1), 1 To 1)
    End If

    captionNo = 0
    For r = 1 To UBound(data, 1)
        key = TrackKey(data, r, colTrackNum, colTrack)
        If Len(key) = 0 Then
            ' Row we could not classify: leave whatever was there untouched.
            outNums(r, 1) = CellValue(data, r, colTrackNum)
            If RESEQUENCE_CAPTION_NUMBERS And colCaption > 0 Then
                outCaptions(r, 1) = CellValue(data, r, colCaption)
            End If
        Else
            idx = IndexOfKey(keys, trackCount, key)
            outNums(r, 1) = newNumber(idx)
            If CellText(data, r, colTrackNum) <> CStr(newNumber(idx)) Then
                changed = changed + 1
            End If
            If RESEQUENCE_CAPTION_NUMBERS And colCaption > 0 Then
                captionNo = captionNo + 1
                outCaptions(r, 1) = captionNo
            End If
        End If
    Next r

    '--- Write back ---------------------------------------------------------
    prevScreen = Application.ScreenUpdating
    prevEvents = Application.EnableEvents
    prevCalc = Application.Calculation
    stateSaved = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    SaveBackup ws, firstRow, lastRow, colTrackNum, colCaption

    ws.Cells(firstRow, colTrackNum).Resize(UBound(data, 1), 1).Value = outNums
    If RESEQUENCE_CAPTION_NUMBERS And colCaption > 0 Then
        ws.Cells(firstRow, colCaption).Resize(UBound(data, 1), 1).Value = outCaptions
    End If

    Application.Calculation = prevCalc
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = prevScreen
    stateSaved = False

    If SHOW_SUMMARY Then
        MsgBox BuildSummary(ws, trackCount, changed, skipped, order, labels, _
                            oldLabel, minStart, hasStart), _
               vbInformation, "Renumber tracks"
    End If
    Exit Sub

Fail:
    If stateSaved Then
        On Error Resume Next
        Application.Calculation = prevCalc
        Application.EnableEvents = prevEvents
        Application.ScreenUpdating = prevScreen
        On Error GoTo 0
    End If
    MsgBox "Renumbering failed: " & Err.Description & " (error " & Err.Number & ").", _
           vbCritical, "Renumber tracks"
End Sub

'--- Track identity -----------------------------------------------------------

' Returns a stable key identifying the track a row belongs to, or "" when the
' row carries neither a track number nor a track name.
Private Function TrackKey(data As Variant, r As Long, _
                          colTrackNum As Long, colTrack As Long) As String
    Dim num As String, name As String

    num = CellText(data, r, colTrackNum)
    name = LCase$(CellText(data, r, colTrack))

    If LCase$(GROUP_BY) = "track" Then
        If Len(name) > 0 Then TrackKey = "N|" & name
        Exit Function
    End If

    If Len(num) > 0 Then
        TrackKey = "#|" & num
    ElseIf Len(name) > 0 Then
        TrackKey = "N|" & name
    End If
End Function

' Human-readable name for a track, used only in the summary message.
Private Function TrackLabel(data As Variant, r As Long, _
                            colTrackNum As Long, colTrack As Long) As String
    TrackLabel = CellText(data, r, colTrack)
    If Len(TrackLabel) = 0 Then TrackLabel = "(unnamed track)"
End Function

Private Function IndexOfKey(keys() As String, count As Long, key As String) As Long
    Dim i As Long
    For i = 1 To count
        If keys(i) = key Then
            IndexOfKey = i
            Exit Function
        End If
    Next i
End Function

'--- Ordering -----------------------------------------------------------------

' Returns track indices ordered by earliest caption. Tracks whose captions have
' no usable start time sort last, in the order they first appear on the sheet.
Private Function SortedTrackOrder(count As Long, minStart() As Double, _
                                  hasStart() As Boolean, firstSeen() As Long) As Long()
    Dim order() As Long
    Dim i As Long, j As Long, cur As Long

    ReDim order(1 To count)
    For i = 1 To count
        order(i) = i
    Next i

    ' Insertion sort: track counts are small and this keeps ties stable.
    For i = 2 To count
        cur = order(i)
        j = i - 1
        Do While j >= 1
            If TrackComesFirst(cur, order(j), minStart, hasStart, firstSeen) Then
                order(j + 1) = order(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        order(j + 1) = cur
    Next i

    SortedTrackOrder = order
End Function

Private Function TrackComesFirst(a As Long, b As Long, minStart() As Double, _
                                 hasStart() As Boolean, firstSeen() As Long) As Boolean
    If hasStart(a) <> hasStart(b) Then
        TrackComesFirst = hasStart(a)
    ElseIf hasStart(a) And minStart(a) <> minStart(b) Then
        TrackComesFirst = (minStart(a) < minStart(b))
    Else
        TrackComesFirst = (firstSeen(a) < firstSeen(b))
    End If
End Function

'--- Undo ---------------------------------------------------------------------

Private Sub SaveBackup(ws As Worksheet, firstRow As Long, lastRow As Long, _
                       colTrackNum As Long, colCaption As Long)
    Dim bk As Worksheet

    On Error Resume Next
    Set bk = ws.Parent.Worksheets(BACKUP_SHEET)
    On Error GoTo 0

    If bk Is Nothing Then
        Set bk = ws.Parent.Worksheets.Add(After:=ws.Parent.Worksheets(ws.Parent.Worksheets.count))
        bk.Name = BACKUP_SHEET
    End If

    bk.Visible = xlSheetVisible
    bk.Cells.Clear
    bk.Range("A1").Value = "Backup written by modTrackNumbering - do not edit"
    bk.Range("A2").Value = "Sheet"
    bk.Range("B2").Value = ws.Name
    bk.Range("A3").Value = "When"
    bk.Range("B3").Value = Now
    bk.Range("A4").Value = "FirstRow"
    bk.Range("B4").Value = firstRow
    bk.Range("A5").Value = "TrackNumberColumn"
    bk.Range("B5").Value = colTrackNum
    bk.Range("A6").Value = "CaptionNumberColumn"
    bk.Range("B6").Value = IIf(RESEQUENCE_CAPTION_NUMBERS, colCaption, 0)
    bk.Range("A7").Value = "RowCount"
    bk.Range("B7").Value = lastRow - firstRow + 1

    bk.Range("D1").Value = "OldTrackNumber"
    bk.Range("D2").Resize(lastRow - firstRow + 1, 1).Value = _
        ws.Cells(firstRow, colTrackNum).Resize(lastRow - firstRow + 1, 1).Value

    If RESEQUENCE_CAPTION_NUMBERS And colCaption > 0 Then
        bk.Range("E1").Value = "OldCaptionNumber"
        bk.Range("E2").Resize(lastRow - firstRow + 1, 1).Value = _
            ws.Cells(firstRow, colCaption).Resize(lastRow - firstRow + 1, 1).Value
    End If

    bk.Visible = xlSheetVeryHidden
End Sub

Public Sub UndoLastTrackRenumber()
    Dim bk As Worksheet, ws As Worksheet
    Dim firstRow As Long, rowCount As Long, colTrackNum As Long, colCaption As Long

    On Error Resume Next
    Set bk = ActiveWorkbook.Worksheets(BACKUP_SHEET)
    On Error GoTo 0

    If bk Is Nothing Then
        MsgBox "No renumbering backup was found in this workbook.", _
               vbExclamation, "Undo renumber"
        Exit Sub
    End If

    On Error Resume Next
    Set ws = ActiveWorkbook.Worksheets(CStr(bk.Range("B2").Value))
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "The backup refers to a sheet named """ & bk.Range("B2").Value & _
               """, which no longer exists.", vbExclamation, "Undo renumber"
        Exit Sub
    End If

    firstRow = CLng(bk.Range("B4").Value)
    colTrackNum = CLng(bk.Range("B5").Value)
    colCaption = CLng(bk.Range("B6").Value)
    rowCount = CLng(bk.Range("B7").Value)

    If rowCount < 1 Then
        MsgBox "The backup is empty.", vbExclamation, "Undo renumber"
        Exit Sub
    End If

    If MsgBox("Restore the track numbers on """ & ws.Name & """ as they were on " & _
              Format$(bk.Range("B3").Value, "yyyy-mm-dd hh:nn:ss") & "?", _
              vbQuestion Or vbYesNo, "Undo renumber") <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    ws.Cells(firstRow, colTrackNum).Resize(rowCount, 1).Value = _
        bk.Range("D2").Resize(rowCount, 1).Value
    If colCaption > 0 Then
        ws.Cells(firstRow, colCaption).Resize(rowCount, 1).Value = _
            bk.Range("E2").Resize(rowCount, 1).Value
    End If
    Application.ScreenUpdating = True

    MsgBox rowCount & " row(s) restored on """ & ws.Name & """.", _
           vbInformation, "Undo renumber"
End Sub

'--- Sheet helpers ------------------------------------------------------------

Private Function HeaderLastColumn(ws As Worksheet) As Long
    HeaderLastColumn = ws.Cells(HEADER_ROW, ws.Columns.count).End(xlToLeft).Column
    If HeaderLastColumn < 1 Then HeaderLastColumn = 1
End Function

Private Function LastDataRow(ws As Worksheet, ParamArray cols() As Variant) As Long
    Dim i As Long, r As Long, c As Long

    LastDataRow = HEADER_ROW
    For i = LBound(cols) To UBound(cols)
        c = CLng(cols(i))
        If c > 0 Then
            r = ws.Cells(ws.Rows.count, c).End(xlUp).Row
            If r > LastDataRow Then LastDataRow = r
        End If
    Next i
End Function

' Finds a column by header text. "candidates" is a "|"-separated list of
' normalised header names, tried in order.
Private Function FindColumn(ws As Worksheet, lastCol As Long, candidates As String) As Long
    Dim wanted() As String
    Dim i As Long, c As Long
    Dim header As String

    wanted = Split(candidates, "|")

    For i = LBound(wanted) To UBound(wanted)
        For c = 1 To lastCol
            header = NormalizeHeader(CStr(ws.Cells(HEADER_ROW, c).Value))
            If header = wanted(i) Then
                FindColumn = c
                Exit Function
            End If
        Next c
    Next i
End Function

' Lower-cases a header and strips everything that is not a letter or digit, so
' "StartTime(s)", "Start Time (s)" and "start_time_s" all collapse to the same
' comparable string.
Private Function NormalizeHeader(text As String) As String
    Dim i As Long, ch As String, out As String

    For i = 1 To Len(text)
        ch = LCase$(Mid$(text, i, 1))
        If (ch >= "a" And ch <= "z") Or (ch >= "0" And ch <= "9") Then out = out & ch
    Next i
    NormalizeHeader = out
End Function

'--- Value helpers ------------------------------------------------------------

Private Function CellValue(data As Variant, r As Long, c As Long) As Variant
    If c < 1 Or c > UBound(data, 2) Then
        CellValue = Empty
    Else
        CellValue = data(r, c)
    End If
End Function

Private Function CellText(data As Variant, r As Long, c As Long) As String
    Dim v As Variant

    v = CellValue(data, r, c)
    If IsError(v) Or IsEmpty(v) Then
        CellText = ""
    Else
        CellText = Trim$(CStr(v))
    End If
End Function

Private Function IsNumericValue(v As Variant) As Boolean
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        If Len(Trim$(CStr(v))) = 0 Then Exit Function
    End If
    IsNumericValue = IsNumeric(v)
End Function

'--- Reporting ----------------------------------------------------------------

Private Function BuildSummary(ws As Worksheet, trackCount As Long, changed As Long, _
                              skipped As Long, order() As Long, labels() As String, _
                              oldLabel() As String, minStart() As Double, _
                              hasStart() As Boolean) As String
    Dim s As String
    Dim i As Long, idx As Long
    Dim shown As Long
    Const MAX_LINES As Long = 40

    s = "Sheet: " & ws.Name & vbCrLf & _
        trackCount & " track(s) renumbered by first caption." & vbCrLf & _
        changed & " caption row(s) had their TrackNumber changed."
    If skipped > 0 Then
        s = s & vbCrLf & skipped & " row(s) were skipped (no track number or name)."
    End If
    s = s & vbCrLf & vbCrLf & "old  ->  new   first caption   track" & vbCrLf

    For i = 1 To trackCount
        If shown >= MAX_LINES Then
            s = s & "... and " & (trackCount - shown) & " more." & vbCrLf
            Exit For
        End If
        idx = order(i)
        s = s & PadRight(oldLabel(idx), 4) & " ->  " & PadRight(CStr(i), 5) & "  " & _
            PadRight(IIf(hasStart(idx), Format$(minStart(idx), "0.000") & "s", "(no time)"), 12) & _
            "  " & labels(idx) & vbCrLf
        shown = shown + 1
    Next i

    BuildSummary = s
End Function

Private Function PadRight(text As String, width As Long) As String
    If Len(text) >= width Then
        PadRight = text
    Else
        PadRight = text & Space$(width - Len(text))
    End If
End Function
