#Requires AutoHotkey v2.0
/*
    ExcelBridge2 - reads the "Soup EE" annotation schema out of a workbook that
    is already open in Excel.

    Deliberately a SEPARATE copy of lib\ExcelBridge.ahk rather than a shared
    file. The two schemas differ in header names, column order and time format,
    and the original filler is trusted working software - a change made for one
    schema must not be able to break the other.

    Schema (verbatim, in order):
        Caption Number, TrackNumber, Track, TrackDescription, StartTime(s),
        EndTime(s), SourceVisibility, SourceDescription, Prominence,
        AnnotationText
*/

class ExcelBridge2 {
    app        := 0
    book       := 0
    sheet      := 0
    columns    := Map()   ; header text -> Excel column number
    rows       := []      ; row records, in sheet order
    tracks     := []      ; distinct track numbers, ascending
    candidates := []      ; open workbooks found across all Excel instances

    /*  The ten headers, in sheet order. The first one doubles as the anchor
        that marks where a header block begins - see FindHeaderBlock.  */
    static Schema := ["Caption Number", "TrackNumber", "Track", "TrackDescription"
                    , "StartTime(s)", "EndTime(s)", "SourceVisibility"
                    , "SourceDescription", "Prominence", "AnnotationText"]

    /*  Only these must be present for a sheet to count as annotations.

        The workbooks vary: some carry all ten columns, others omit
        SourceDescription entirely. Demanding the full set made the tool reject
        perfectly good sheets. Anything outside this list is optional - when a
        sheet lacks it, the field simply reports itself as absent and its button
        stays disabled.

        "Caption Number" is also the block anchor, so it cannot be optional.  */
    static Required := ["Caption Number", "Track", "StartTime(s)", "EndTime(s)"
                      , "AnnotationText"]

    /*  Property name used on a row record, per header.  */
    static Fields := Map(
        "Caption Number"    , "captionNo",
        "TrackNumber"       , "trackNo",
        "Track"             , "track",
        "TrackDescription"  , "trackDesc",
        "StartTime(s)"      , "startTime",
        "EndTime(s)"        , "endTime",
        "SourceVisibility"  , "sourceVis",
        "SourceDescription" , "sourceDesc",
        "Prominence"        , "prominence",
        "AnnotationText"    , "annotation")

    ; ------------------------------------------------------------ workbooks --

    /*  Finds open workbooks and returns their names.

        Does not rely on ComObjActive alone: that returns whichever instance
        registered first, which is regularly a leftover invisible instance
        holding no workbooks, leaving the real one unreachable. Every Excel
        window is asked for its own object model instead.  */
    Attach() {
        this.candidates := []
        seen := Map()

        for hwnd in WinGetList("ahk_class XLMAIN") {
            app := ExcelAppFromWindow2(hwnd)
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

    CollectFrom(app, seen) {
        try {
            for book in app.Workbooks {
                name := book.Name
                if (SubStr(name, 1, 1) = "~")          ; Excel lock/temp files
                    continue
                key := name
                try key := book.FullName               ; may be a OneDrive URL
                if seen.Has(key)
                    continue
                seen[key] := true
                this.candidates.Push({ name: name, book: book, app: app })
            }
        }
    }

    BindWorkbook(name) {
        for candidate in this.candidates {
            if (candidate.name = name) {
                this.book := candidate.book
                this.app  := candidate.app
                return
            }
        }
        throw Error("Workbook '" name "' is no longer open.")
    }

    ; --------------------------------------------------------------- sheets --

    /*  Names of every sheet in the bound workbook that carries the schema.
        A workbook typically holds several: "Completion" and "Refinement " both
        match, while a "Summary" sheet of Field/Value pairs does not.  */
    MatchingSheets() {
        matches := []
        for sheet in this.book.Worksheets {
            if this.FindHeaderBlock(sheet)
                matches.Push(sheet.Name)     ; verbatim - names can end in a space
        }
        return matches
    }

    BindSheet(name) {
        for sheet in this.book.Worksheets {
            if (sheet.Name == name) {        ; == : case sensitive, keeps a trailing space honest
                columns := this.FindHeaderBlock(sheet)
                if !columns
                    throw Error("Sheet '" name "' does not carry the expected columns.")
                this.sheet := sheet
                this.columns := columns
                return
            }
        }
        throw Error("Sheet '" name "' is no longer in this workbook.")
    }

    /*  Locates the schema on a sheet and returns header -> column number, or 0.

        The sheets in play repeat the same ten headers up to three times across
        the columns: the real data, then a variant, then a block of TRUE/FALSE
        change flags. Mapping headers naively (last one wins, as the original
        bridge does) would bind Track to a column of booleans and type FALSE
        into the website.

        So: the block starts at the first occurrence of the anchor header and
        ends just before the anchor appears again. Only that window is searched,
        and within it the first occurrence of each header wins. Extra or
        interleaved columns inside the block are tolerated.  */
    FindHeaderBlock(sheet) {
        anchor := ExcelBridge2.Schema[1]

        used := 0
        try used := sheet.UsedRange
        if !used
            return 0

        firstCol := used.Column
        lastCol  := used.Column + used.Columns.Count - 1
        headerRow := used.Row
        if (lastCol < firstCol)
            return 0

        values := 0
        try values := sheet.Range(sheet.Cells(headerRow, firstCol)
                                , sheet.Cells(headerRow, lastCol)).Value2
        if !IsObject(values)
            return 0

        count := 0
        try count := values.MaxIndex(2)
        if (!count || count < 1)
            return 0

        ; Header text per column, trimmed for comparison.
        headers := []
        loop count {
            text := ""
            try text := Trim(String(values[1, A_Index]))
            headers.Push(text)
        }

        ; Where does the block start, and where does the next one begin?
        start := 0, stop := headers.Length
        loop headers.Length {
            if (headers[A_Index] = anchor) {
                if !start
                    start := A_Index
                else {
                    stop := A_Index - 1
                    break
                }
            }
        }
        if !start
            return 0

        ; First occurrence wins, searching only inside the block.
        columns := Map()
        index := start
        while (index <= stop) {
            header := headers[index]
            if (header != "")
                for want in ExcelBridge2.Schema
                    if (header = want && !columns.Has(want)) {
                        ; Comparisons above are intentionally case-insensitive,
                        ; but Map keys are case-sensitive. Keep the canonical
                        ; schema spelling so Fields[header] and HasColumn() use
                        ; the same key even when Excel says, for example,
                        ; "Annotationtext" instead of "AnnotationText".
                        columns[want] := index + firstCol - 1
                        break
                    }
            index++
        }

        for want in ExcelBridge2.Required
            if !columns.Has(want)
                return 0

        return columns
    }

    /*  Does the bound sheet actually carry this column?
        Distinguishes "the cell is empty" from "there is no such column".  */
    HasColumn(header) {
        return this.columns.Has(header)
    }

    ; ----------------------------------------------------------------- rows --

    /*  Reads every data row under the header, in one COM round trip.  */
    Load() {
        if !this.sheet
            throw Error("No sheet is bound yet.")

        used := this.sheet.UsedRange
        headerRow := used.Row
        firstCol  := used.Column
        data := used.Value2

        rowCount := 0
        try rowCount := data.MaxIndex(1)
        if (!rowCount || rowCount < 2)
            throw Error("Sheet '" this.sheet.Name "' has a header row but no annotations under it.")

        this.rows := []
        trackSeen := Map()

        loop rowCount - 1 {
            r := A_Index + 1
            record := { excelRow: r + headerRow - 1 }

            ; Every field must exist even when its column does not, or reading
            ; an absent one later throws instead of simply coming back empty.
            for header in ExcelBridge2.Schema
                record.%ExcelBridge2.Fields[header]% := ""

            for header, column in this.columns {
                field := ExcelBridge2.Fields[header]
                raw := ""
                try raw := data[r, column - firstCol + 1]
                record.%field% := (raw = "") ? "" : Trim(String(raw))
            }

            ; An entirely blank row is trailing noise, not an annotation.
            if (record.annotation = "" && record.track = ""
                && record.startTime = "" && record.endTime = "")
                continue

            record.captionNo := NormalizeCount(record.captionNo)
            record.trackNo   := NormalizeCount(record.trackNo)
            record.startTime := FormatSeconds(record.startTime)
            record.endTime   := FormatSeconds(record.endTime)

            this.rows.Push(record)
            if (record.trackNo != "")
                trackSeen[record.trackNo] := true
        }

        if !this.rows.Length
            throw Error("No annotation rows found on sheet '" this.sheet.Name "'.")

        this.tracks := []
        for trackNo, _ in trackSeen
            this.tracks.Push(trackNo)
        SortNumbers2(this.tracks)
    }

    /*  Re-reads one row so edits made in Excel appear as you navigate.  */
    Refresh(index) {
        record := this.rows[index]

        least := 0, most := 0
        for header, column in this.columns {
            if (!least || column < least)
                least := column
            if (column > most)
                most := column
        }

        values := 0
        try values := this.sheet.Range(this.sheet.Cells(record.excelRow, least)
                                     , this.sheet.Cells(record.excelRow, most)).Value2
        if !IsObject(values)
            return record

        for header, column in this.columns {
            field := ExcelBridge2.Fields[header]
            raw := ""
            try raw := values[1, column - least + 1]
            record.%field% := (raw = "") ? "" : Trim(String(raw))
        }

        record.captionNo := NormalizeCount(record.captionNo)
        record.trackNo   := NormalizeCount(record.trackNo)
        record.startTime := FormatSeconds(record.startTime)
        record.endTime   := FormatSeconds(record.endTime)
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
}

/*  Returns the Excel Application object belonging to one Excel window.

    Each Excel window owns an "EXCEL7" pane exposing the native object model
    through AccessibleObjectFromWindow(OBJID_NATIVEOM) - the only way to reach a
    specific instance when several are running.  */
ExcelAppFromWindow2(hwnd) {
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
        window := ComValue(9, pAcc)
        return window.Application
    }
    return 0
}

/*  Seconds, formatted the way the sheet displays them: three decimals.
    Formatting the number rather than passing it through also removes Excel's
    floating point noise, which otherwise turns 4.448 into 4.4480000000000004. */
FormatSeconds(value) {
    if (value = "" || !IsNumber(value))
        return value = "" ? "" : value      ; keep non-numeric text as written
    return Format("{:.3f}", value + 0)
}

/*  Splits seconds into whole minutes / seconds / milliseconds, because the
    website takes a time in three boxes while this schema stores it in one.
    Rounds to milliseconds FIRST, so 59.9996 becomes 1 min 0 s 0 ms rather than
    an impossible 60 in the seconds box.

    Text that is not a number splits to nothing at all - the seconds box still
    shows it verbatim, so a malformed cell stays visible instead of being
    quietly turned into 0 min 0 s 0 ms.  */
SplitTime2(seconds) {
    if (seconds = "" || !IsNumber(seconds))
        return { min: "", sec: "", ms: "" }
    totalMs := Round(seconds * 1000)
    if (totalMs < 0)
        totalMs := 0
    return { min: totalMs // 60000
           , sec: Mod(totalMs // 1000, 60)
           , ms : Mod(totalMs, 1000) }
}

/*  Caption and track numbers arrive from Excel as doubles: 4 comes back as 4.0
    and would be typed as "4.0". Whole numbers collapse to integers.  */
NormalizeCount(value) {
    if (value = "" || !IsNumber(value))
        return value
    value += 0
    return (value = Round(value)) ? Integer(value) : value
}

Join2(arr, sep) {
    out := ""
    for item in arr
        out .= (out = "" ? "" : sep) item
    return out
}

/*  Ascending numeric sort of a plain array, in place.  */
SortNumbers2(arr) {
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
