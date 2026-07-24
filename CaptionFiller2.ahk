#Requires AutoHotkey v2.0
#SingleInstance Force
/*
    AutoNyx Caption Filler II
    -------------------------
    A floating guide for the "Soup EE" column schema:

        Caption Number, TrackNumber, Track, TrackDescription, StartTime(s),
        EndTime(s), SourceVisibility, SourceDescription, Prominence,
        AnnotationText

    All ten are shown. Six can be typed into the website with one click:
    Track, TrackDescription, StartTime(s), EndTime(s), SourceDescription and
    AnnotationText. The other four are read-only context.

    How you use it:
        1. click the target box on the website  (the browser keeps focus)
        2. click the matching button here       (the value is typed into it)

    The guide never clicks anything in the browser, never presses Tab, and
    never submits the form. It only types into the field you focused.

    This is a SEPARATE tool from CaptionFiller.ahk, which serves a different
    schema. Neither shares code with the other, on purpose.
*/

#Include lib\ExcelBridge2.ahk

; ---------------------------------------------------------------- constants --

APP_TITLE := "AutoNyx Caption Filler II"
INI       := A_ScriptDir "\CaptionFiller2.ini"

/*  Glyphs via Chr() so this file stays pure ASCII: AutoHotkey reads a script
    with no UTF-8 BOM in the system codepage, which silently turns pasted
    arrows and ticks into "?".
    Named Sym, not G - a class name collides with any same-named variable,
    and identifiers here are case-insensitive.  */
class Sym {
    static PREV := Chr(0x25C0)     ; left triangle
    static NEXT := Chr(0x25B6)     ; right triangle
    static SENT := Chr(0x2713)     ; check mark
    static ELL  := Chr(0x2026)     ; ellipsis
}

/*  Every column, in sheet order. "fill" marks the ones with a button.
    field must match ExcelBridge2.Fields.  */
FIELDS := [
    { field: "captionNo" , label: "Caption Number"    , fill: false },
    { field: "trackNo"   , label: "TrackNumber"       , fill: false },
    { field: "track"     , label: "Track"             , fill: true  },
    { field: "trackDesc" , label: "TrackDescription"  , fill: true  },
    { field: "startTime" , label: "StartTime(s)"      , fill: true  },
    { field: "endTime"   , label: "EndTime(s)"        , fill: true  },
    { field: "sourceVis" , label: "SourceVisibility"  , fill: false },
    { field: "sourceDesc", label: "SourceDescription" , fill: true  },
    { field: "prominence", label: "Prominence"        , fill: false },
    { field: "annotation", label: "AnnotationText"    , fill: true  } ]

; ------------------------------------------------------------------- state --

bridge   := ExcelBridge2()
index    := 1                ; 1-based position in bridge.rows
values   := Map()            ; field -> string currently shown
sent     := Map()            ; field -> true once dispensed
paused   := false
guiHwnd  := 0
lastHwnd := 0                ; last foreground window that was not the guide
lastCtrl := 0                ; and the control focused inside it
ui       := Map()
bound    := { workbook: "", sheet: "" }

; ------------------------------------------------------------------ startup --

Main()

Main() {
    global bridge, index, bound

    try {
        bridge.Attach()
        choice := PickSource()
        if !choice
            ExitApp
        bridge.BindWorkbook(choice.workbook)
        bridge.BindSheet(choice.sheet)
        bridge.Load()
        bound := choice
    } catch as e {
        MsgBox(e.Message, APP_TITLE, "Iconx")
        ExitApp
    }

    BuildGui()

    remembered := Integer(IniRead(INI, "state", "index", "1"))
    index := (remembered >= 1 && remembered <= bridge.rows.Length) ? remembered : 1

    ShowCaption()
    SetTimer(TrackFocus, 250)
}

/*  Every workbook/sheet pair carrying the schema. A workbook holds several
    such sheets, and several workbooks may be open, so this is a flat list.  */
ScanSources() {
    global bridge
    sources := []
    for candidate in bridge.candidates {
        try {
            bridge.BindWorkbook(candidate.name)
            for sheetName in bridge.MatchingSheets()
                sources.Push({ workbook: candidate.name, sheet: sheetName })
        }
    }
    return sources
}

/*  Binds silently when only one sheet anywhere matches; asks otherwise.  */
PickSource() {
    global bridge

    sources := ScanSources()
    if !sources.Length {
        MsgBox("None of the open workbooks contain the expected columns.`n`n"
             . "Expected, in any position:`n    "
             . Join2(ExcelBridge2.Schema, "`n    ")
             . "`n`nOpen the right workbook in Excel and start this tool again."
             , APP_TITLE, "Iconx")
        return 0
    }

    ; An argument names the workbook outright: CaptionFiller2.ahk "Soup EE.xlsx"
    if A_Args.Length {
        for source in sources
            if (source.workbook = A_Args[1])
                return source
    }

    if (sources.Length = 1)
        return sources[1]

    labels := []
    for source in sources
        labels.Push(source.workbook "   " Sym.NEXT "   " source.sheet)

    preferred := IniRead(INI, "state", "source", "")
    picked := 0
    chooser := Gui("+AlwaysOnTop -MinimizeBox", APP_TITLE)
    chooser.SetFont("s10", "Segoe UI")
    chooser.Add("Text",, "Which sheet holds the annotations?")
    list := chooser.Add("ListBox", "w460 r" Min(12, labels.Length), labels)
    list.Choose(1)
    for i, label in labels
        if (label = preferred)
            list.Choose(i)
    ok := chooser.Add("Button", "w120 Default", "Use this")
    cancel := chooser.Add("Button", "x+10 yp w120", "Cancel")
    ok.OnEvent("Click", (*) => (picked := list.Value, chooser.Destroy()))
    cancel.OnEvent("Click", (*) => chooser.Destroy())
    chooser.OnEvent("Close", (*) => chooser.Destroy())
    chooser.Show()
    WinWaitClose("ahk_id " chooser.Hwnd)

    if !picked
        return 0
    IniWrite(labels[picked], INI, "state", "source")
    return sources[picked]
}

; --------------------------------------------------------------------- view --

BuildGui() {
    global ui, guiHwnd, bound

    win := Gui("AlwaysOnTop -MinimizeBox -MaximizeBox +E0x08000000 -DPIScale"
             , APP_TITLE " - " bound.workbook " [" bound.sheet "]")
    win.SetFont("s9", "Segoe UI")
    win.MarginX := 10, win.MarginY := 8
    ui["gui"] := win

    ; --- orientation and navigation ---
    ui["trackText"] := win.Add("Text", "w400 h18", "Track -")
    btn := win.Add("Button", "x+6 yp-2 w76 h24", Sym.PREV " Track")
    btn.OnEvent("Click", (*) => GoTrack(-1))
    ui["prevTrack"] := btn
    btn := win.Add("Button", "x+4 yp w76 h24", "Track " Sym.NEXT)
    btn.OnEvent("Click", (*) => GoTrack(1))
    ui["nextTrack"] := btn

    ui["captionText"] := win.Add("Text", "xm y+6 w400 h18", "Caption -")
    btn := win.Add("Button", "x+6 yp-2 w76 h24", Sym.PREV " Cap")
    btn.OnEvent("Click", (*) => GoCaption(-1))
    ui["prevCaption"] := btn
    btn := win.Add("Button", "x+4 yp w76 h24", "Cap " Sym.NEXT)
    btn.OnEvent("Click", (*) => GoCaption(1))
    ui["nextCaption"] := btn

    win.Add("Text", "xm y+8 w578 h1 0x10")

    ; --- the ten columns, in sheet order ---
    for entry in FIELDS {
        key := entry.field
        if entry.fill {
            btn := win.Add("Button", "xm y+5 w170 h26", "   " entry.label)
            btn.OnEvent("Click", Fill.Bind(key))
            ui["btn_" key] := btn
            ui["value_" key] := win.Add("Edit", "x+8 yp+2 w400 h22 ReadOnly -E0x200")
        } else {
            ; Read-only context: shown, but nothing to click.
            win.Add("Text", "xm y+5 w170 h22 +0x200", "      " entry.label)
            ui["value_" key] := win.Add("Edit", "x+8 yp w400 h22 ReadOnly -E0x200 +Background" . "F0F0F0")
        }
    }

    win.Add("Text", "xm y+10 w578 h1 0x10")

    ; --- status and options ---
    ui["status"] := win.Add("Text", "xm y+6 w578 h18", "Ready.")

    ui["replace"] := win.Add("Checkbox", "xm y+6 Checked", "Replace field contents")
    ui["follow"]  := win.Add("Checkbox", "x+14 yp Checked", "Follow in Excel")

    btn := win.Add("Button", "xm y+8 w96 h26", "Reload")
    btn.OnEvent("Click", (*) => ReloadSheet())
    ui["pause"] := win.Add("Button", "x+6 yp w96 h26", "Pause")
    ui["pause"].OnEvent("Click", (*) => TogglePause())

    ui["replace"].Value := Integer(IniRead(INI, "options", "replace", "1"))
    ui["follow"].Value  := Integer(IniRead(INI, "options", "follow", "1"))

    x := IniRead(INI, "window", "x", "")
    y := IniRead(INI, "window", "y", "")
    win.OnEvent("Close", (*) => Shutdown())
    if (x != "" && y != "")
        win.Show("NoActivate x" x " y" y)
    else
        win.Show("NoActivate")
    guiHwnd := win.Hwnd
}

ShowCaption() {
    global bridge, index, values, sent, ui

    record := bridge.Refresh(index)

    values := Map()
    for entry in FIELDS
        values[entry.field] := record.%entry.field%

    sent := Map()

    ui["trackText"].Value := (record.trackNo = "")
        ? "Track    -"
        : Format("Track    {1} / {2}    (TrackNumber {3})"
            , TrackPosition(record.trackNo), bridge.tracks.Length, record.trackNo)
    ui["captionText"].Value := Format("Caption  {1} / {2}    (Caption Number {3}, sheet row {4})"
        , index, bridge.rows.Length, record.captionNo, record.excelRow)

    for entry in FIELDS {
        value := values[entry.field]
        ; A column the sheet does not have at all reads differently from a
        ; column whose cell happens to be blank.
        present := bridge.HasColumn(entry.label)
        ui["value_" entry.field].Value := present ? value : "(not in this sheet)"
        if entry.fill
            ui["btn_" entry.field].Enabled := (present && value != "")
    }

    ui["prevCaption"].Enabled := (index > 1)
    ui["nextCaption"].Enabled := (index < bridge.rows.Length)
    ui["prevTrack"].Enabled := (AdjacentTrack(-1) != 0)
    ui["nextTrack"].Enabled := (AdjacentTrack(1) != 0)

    MarkNext()
    Status(Shorten(record.annotation, 80))
    if ui["follow"].Value
        bridge.ShowRow(index)
    IniWrite(index, INI, "state", "index")
}

/*  Leading glyph per fill button: sent, next up, or neither.  */
MarkNext() {
    global ui, sent, values

    nextKey := ""
    for entry in FIELDS {
        if (entry.fill && values[entry.field] != "" && !sent.Has(entry.field)) {
            nextKey := entry.field
            break
        }
    }
    for entry in FIELDS {
        if !entry.fill
            continue
        mark := sent.Has(entry.field) ? Sym.SENT " "
              : (entry.field = nextKey ? Sym.NEXT " " : "   ")
        ui["btn_" entry.field].Text := mark entry.label
    }
}

; ---------------------------------------------------------------- behaviour --

/*  Remembers where a value should go: the last window in front that was not
    this guide, plus the control focused inside it.  */
TrackFocus() {
    global guiHwnd, lastHwnd, lastCtrl
    active := WinExist("A")
    if (!active || active = guiHwnd)
        return
    lastHwnd := active
    try lastCtrl := ControlGetFocus(active)
    catch
        lastCtrl := 0
}

Fill(key, *) {
    global values, sent, ui, paused, lastHwnd, lastCtrl

    ; SendText returns before the receiving window has processed the keys.
    ; Without this guard a quick second click can land its Ctrl+A in the middle
    ; of the previous value and shred both.
    static busy := false
    if busy {
        Status("Still typing - one moment.")
        return
    }
    if paused {
        Status("Paused - press Resume to type again.")
        return
    }

    value := values.Has(key) ? values[key] : ""
    if (value = "") {
        Status("Nothing to send for that field.")
        return
    }
    if !lastHwnd {
        Status("Click the website field first, then click here.")
        return
    }
    ; Refuse to type into Excel: clicking a cell to check something and then a
    ; fill button would otherwise overwrite the annotation data itself.
    try {
        if (WinGetProcessName(lastHwnd) = "EXCEL.EXE") {
            Status("That would type into Excel. Click the website field first.")
            return
        }
    }

    ; Safety net. With WS_EX_NOACTIVATE the guide never took focus and this does
    ; nothing; without it, this puts focus back where the value belongs.
    if (WinExist("A") != lastHwnd) {
        try {
            WinActivate(lastHwnd)
            WinWaitActive(lastHwnd,, 1)
            if lastCtrl
                ControlFocus(lastCtrl)
            Sleep 80
        } catch {
            Status("Could not return focus to the field.")
            return
        }
    }

    busy := true
    try {
        if ui["replace"].Value
            Send("^a")
        SendText(value)     ; SendText, never Send: text contains { } ! ^ + #
        Sleep 60 + Min(400, StrLen(value) * 3)
    } finally {
        busy := false
    }

    sent[key] := true
    MarkNext()
    Status("Sent " LabelOf(key) ": " Shorten(value, 60))
}

GoCaption(delta) {
    global index, bridge
    target := index + delta
    if (target < 1 || target > bridge.rows.Length)
        return
    index := target
    ShowCaption()
}

GoTrack(direction) {
    global index
    target := AdjacentTrack(direction)
    if !target
        return
    index := target
    ShowCaption()
}

/*  First row of the next/previous track NUMBER, or 0 if there is none.
    Neither sheet is sorted by track, so this walks the sorted list of distinct
    track numbers and then finds that track's first row in sheet order.  */
AdjacentTrack(direction) {
    global bridge, index

    current := bridge.rows[index].trackNo
    if (current = "" || !bridge.tracks.Length)
        return 0

    wanted := ""
    for trackNo in bridge.tracks {
        if (direction > 0 && trackNo > current) {
            wanted := trackNo
            break
        }
        if (direction < 0 && trackNo < current)
            wanted := trackNo          ; keep the largest still smaller
    }
    if (wanted = "")
        return 0

    for i, record in bridge.rows
        if (record.trackNo = wanted)
            return i
    return 0
}

TrackPosition(trackNo) {
    global bridge
    for i, value in bridge.tracks
        if (value = trackNo)
            return i
    return "-"
}

ReloadSheet() {
    global bridge, index
    try {
        bridge.Load()
        if (index > bridge.rows.Length)
            index := bridge.rows.Length
        ShowCaption()
        Status("Reloaded " bridge.rows.Length " captions.")
    } catch as e {
        MsgBox(e.Message, APP_TITLE, "Iconx")
    }
}

TogglePause() {
    global paused, ui
    paused := !paused
    ui["pause"].Text := paused ? "Resume" : "Pause"
    Status(paused ? "Paused - no typing until you press Resume." : "Ready.")
}

Status(text) {
    global ui
    ui["status"].Value := text
}

LabelOf(key) {
    for entry in FIELDS
        if (entry.field = key)
            return entry.label
    return key
}

Shorten(text, limit) {
    text := StrReplace(StrReplace(text, "`r", " "), "`n", " ")
    return StrLen(text) > limit ? SubStr(text, 1, limit - 1) Sym.ELL : text
}

Shutdown() {
    global ui, index
    try {
        ui["gui"].GetPos(&x, &y)
        IniWrite(x, INI, "window", "x")
        IniWrite(y, INI, "window", "y")
        IniWrite(ui["replace"].Value, INI, "options", "replace")
        IniWrite(ui["follow"].Value, INI, "options", "follow")
        IniWrite(index, INI, "state", "index")
    }
    ExitApp
}
