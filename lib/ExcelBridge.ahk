#Requires AutoHotkey v2.0
/*
    Reads annotation rows out of a workbook that is already open in Excel.

    Columns are located by HEADER TEXT, never by position: the layout is owned
    by $AnnotationColumns in lib\AnnotationWorkbook.ps1, and renaming or
    reordering a column there must not break this tool.
*/

class ExcelBridge {
    app       := 0
    book      := 0
    sheet     := 0
    columns   := Map()   ; header text -> Excel column number
    rows      := []      ; row records, in sheet order
    tracks    := []      ; distinct track numbers, ascending
    colOffset := 0       ; used range may not start at column A

    ; Headers we cannot work without.
    static Required := ["Caption", "Start Time", "End Time"]

    candidates := []     ; every open workbook found, across all Excel instances

    /*  Finds open workbooks and returns their names.

        Deliberately does NOT rely on ComObjActive alone. That returns whichever
        instance registered itself first, which is regularly the wrong one: a
        second Excel instance, or a leftover invisible instance holding no
        workbooks at all, in which case the real workbook is invisible to us.
        Instead every Excel window is asked for its own object model, and
        ComObjActive is consulted last as a bonus (it can see a workbook that
        has no window).  */
    Attach() {
        this.candidates := []
        seen := Map()

        for hwnd in WinGetList("ahk_class XLMAIN") {
            app := ExcelAppFromWindow(hwnd)
            if app
                this.CollectFrom(app, seen)
        }
        try this.CollectFrom(ComObjActive("Excel.Application"), seen)

        if !this.candidates.Length {
            throw Error("No open Excel workbook could be found.`n`n"
                      . "Open your annotation workbook in Excel, then press Reload.`n`n"
                      . "(If Excel is open but empty, or was started as administrator "
                      . "while this script was not, it cannot be reached.)")
        }

        names := []
        for candidate in this.candidates
            names.Push(candidate.name)
        return names
    }

    /*  Adds an instance's workbooks, skipping duplicates and lock files.  */
    CollectFrom(app, seen) {
        try {
            for book in app.Workbooks {
                name := book.Name
                if (SubStr(name, 1, 1) = "~")           ; Excel lock/temp files
                    continue
                key := name
                try key := book.FullName                ; may be a OneDrive URL
                if seen.Has(key)
                    continue
                seen[key] := true
                this.candidates.Push({ name: name, fullName: key, book: book, app: app })
            }
        }
    }

    BindWorkbook(name) {
        for candidate in this.candidates {
            if (candidate.name = name) {
                this.book := candidate.book
                this.app := candidate.app
                this.sheet := candidate.book.ActiveSheet
                return
            }
        }
        throw Error("Workbook '" name "' is no longer open.")
    }

    /*  Fallback when the workbook lives in an Excel instance ComObjActive did
        not hand us: binding by file path works across instances.  */
    BindFile(path) {
        this.book := ComObjGet(path)
        this.sheet := this.book.ActiveSheet
        this.app := this.book.Application
    }

    /*  Reads the header row and every data row beneath it, in one COM call.  */
    Load() {
        used := this.sheet.UsedRange
        rowOffset := used.Row - 1
        this.colOffset := used.Column - 1
        data := used.Value2          ; one round trip; cell-by-cell COM is glacial

        rowCount := 0, colCount := 0
        try {
            rowCount := data.MaxIndex(1)
            colCount := data.MaxIndex(2)
        } catch {
            throw Error("Sheet '" this.sheet.Name "' does not look like a table of annotations.")
        }
        if (rowCount < 2)
            throw Error("Sheet '" this.sheet.Name "' has a header row but no annotations under it.")

        ; --- headers -> Excel column numbers ---
        this.columns := Map()
        loop colCount {
            header := ""
            try header := Trim(String(data[1, A_Index]))
            if (header != "")
                this.columns[header] := A_Index + this.colOffset
        }

        missing := []
        for want in ExcelBridge.Required
            if !this.columns.Has(want)
                missing.Push(want)
        if missing.Length {
            found := ""
            for header, _ in this.columns
                found .= (found = "" ? "" : ", ") header
            throw Error("This sheet is missing: " Join(missing, ", ")
                      . "`n`nHeaders found: " (found = "" ? "(none)" : found)
                      . "`n`nSwitch to the sheet made by the AutoNyx export tools.")
        }

        ; --- rows ---
        this.rows := []
        trackSeen := Map()
        loop rowCount - 1 {
            r := A_Index + 1
            record := {
                excelRow   : r + rowOffset,
                caption    : this.Text(data, r, "Caption"),
                startTime  : this.Num(data, r, "Start Time"),
                endTime    : this.Num(data, r, "End Time"),
                captionNo  : this.Num(data, r, "Caption Number"),
                trackNo    : this.Num(data, r, "Track Number"),
                trackName  : this.Text(data, r, "Track Name"),
                sourceDesc : this.Text(data, r, "Source Description")
            }
            ; A row with no caption and no times is trailing noise, not data.
            if (record.caption = "" && record.startTime = "" && record.endTime = "")
                continue
            this.rows.Push(record)
            if (record.trackNo != "")
                trackSeen[record.trackNo] := true
        }

        if !this.rows.Length
            throw Error("No annotation rows found on sheet '" this.sheet.Name "'.")

        this.tracks := []
        for trackNo, _ in trackSeen
            this.tracks.Push(trackNo)
        SortNumbers(this.tracks)
    }

    /*  Re-reads one row so edits made in Excel appear as you navigate.
        Cheap: a single COM read of that row's slice.  */
    Refresh(index) {
        record := this.rows[index]
        first := this.colOffset + 1
        last  := this.colOffset + this.columns.Count
        for header, col in this.columns {
            if (col > last)
                last := col
        }
        values := this.sheet.Range(this.sheet.Cells(record.excelRow, first)
                                 , this.sheet.Cells(record.excelRow, last)).Value2

        Get(header) {
            if !this.columns.Has(header)
                return ""
            value := ""
            try value := values[1, this.columns[header] - first + 1]
            return (value = "") ? "" : Trim(String(value))
        }
        GetNum(header) {
            return NormalizeNumber(Get(header))
        }

        record.caption    := Get("Caption")
        record.trackName  := Get("Track Name")
        record.sourceDesc := Get("Source Description")
        record.startTime  := GetNum("Start Time")
        record.endTime    := GetNum("End Time")
        record.captionNo  := GetNum("Caption Number")
        record.trackNo    := GetNum("Track Number")
        return record
    }

    /*  Scrolls Excel to a row. Does not pull Excel to the foreground.  */
    ShowRow(index) {
        record := this.rows[index]
        try {
            this.sheet.Activate()
            this.sheet.Range("A" record.excelRow).Select()
            this.app.ActiveWindow.ScrollRow := Max(1, record.excelRow - 6)
        }
    }

    ; --- helpers reading the cached 2-D array --------------------------------

    Text(data, r, header) {
        if !this.columns.Has(header)
            return ""
        value := ""
        try value := data[r, this.columns[header] - this.colOffset]
        return (value = "") ? "" : Trim(String(value))
    }

    Num(data, r, header) {
        return NormalizeNumber(this.Text(data, r, header))
    }
}

/*  Returns the Excel Application object belonging to one Excel window.

    Each Excel window owns an "EXCEL7" pane that exposes the native object model
    through AccessibleObjectFromWindow(OBJID_NATIVEOM). This is the only way to
    reach a specific instance when several are running.  */
ExcelAppFromWindow(hwnd) {
    static OBJID_NATIVEOM := 0xFFFFFFF0

    pane := 0
    try {
        for ctrl in WinGetControls(hwnd) {
            if (SubStr(ctrl, 1, 6) = "EXCEL7") {
                pane := ControlGetHwnd(ctrl, hwnd)
                break
            }
        }
    }
    if !pane
        return 0

    guid := Buffer(16, 0)
    if DllCall("ole32\CLSIDFromString", "WStr", "{00020400-0000-0000-C000-000000000046}", "Ptr", guid) < 0
        return 0

    pAcc := 0
    hr := DllCall("oleacc\AccessibleObjectFromWindow", "Ptr", pane, "UInt", OBJID_NATIVEOM
                , "Ptr", guid, "Ptr*", &pAcc, "Int")
    if (hr != 0 || !pAcc)
        return 0

    try {
        window := ComValue(9, pAcc)     ; wraps the IDispatch we were handed
        return window.Application
    }
    return 0
}

/*  Excel hands back every number as a double, so a caption number arrives as
    1.0 and would be typed into the website as "1.0". Collapse whole numbers to
    integers; genuine decimals (the timestamps) keep their fraction.  */
NormalizeNumber(value) {
    if (value = "" || !IsNumber(value))
        return ""
    value += 0
    return (value = Round(value) && Abs(value) < 9007199254740992) ? Integer(value) : value
}

/*  Splits seconds into whole minutes / seconds / milliseconds.
    Rounds to milliseconds FIRST, so 59.9996 becomes 1 min 0 s 0 ms rather than
    producing an impossible 60 in the seconds box.  */
SplitTime(seconds) {
    if (seconds = "" || !IsNumber(seconds))
        return { min: "", sec: "", ms: "" }
    totalMs := Round(seconds * 1000)
    if (totalMs < 0)
        totalMs := 0
    return { min: totalMs // 60000
           , sec: Mod(totalMs // 1000, 60)
           , ms : Mod(totalMs, 1000) }
}

Join(arr, sep) {
    out := ""
    for item in arr
        out .= (out = "" ? "" : sep) item
    return out
}

/*  Ascending numeric sort of a plain array, in place.  */
SortNumbers(arr) {
    loop arr.Length - 1 {
        i := A_Index
        loop arr.Length - i {
            j := A_Index
            if (arr[j] > arr[j + 1]) {
                tmp := arr[j], arr[j] := arr[j + 1], arr[j + 1] := tmp
            }
        }
    }
    return arr
}
