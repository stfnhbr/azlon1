#Requires AutoHotkey v2.0
#SingleInstance Force
/*
    AutoNyx Caption Filler
    ----------------------
    A floating guide that holds one caption's values and types them into the
    annotation website, one field per click.

    How you use it:
        1. click the target box on the website  (the browser keeps focus)
        2. click the matching button here       (the value is typed into it)

    The guide never clicks anything in the browser, never presses Tab, and
    never submits the form. It only types into the field you focused. That is
    what makes it safe to point at a real work site.

    The window carries WS_EX_NOACTIVATE so clicking it does not steal focus.
    Should that ever fail, every fill also restores the last focused window and
    control first, so the value still lands in the right place.
*/

#Include lib\ExcelBridge.ahk

; ---------------------------------------------------------------- constants --

APP_TITLE := "AutoNyx Caption Filler"
INI       := A_ScriptDir "\CaptionFiller.ini"

/*  Glyphs are built with Chr() so this file stays pure ASCII. AutoHotkey reads
    a script without a UTF-8 BOM using the system codepage, which silently turns
    pasted arrows and ticks into "?" - keeping the source ASCII removes the
    whole class of problem.  */
class Glyph {
    static PREV := Chr(0x25C0)     ; left triangle
    static NEXT := Chr(0x25B6)     ; right triangle
    static SENT := Chr(0x2713)     ; check mark
    static ELL  := Chr(0x2026)     ; ellipsis
}

; The eight fillable fields, in the order the guide walks them.
FIELDS := [
    { key: "caption",   label: "Caption",            wide: true  },
    { key: "startMin",  label: "Start Min",          wide: false },
    { key: "startSec",  label: "Start Sec",          wide: false },
    { key: "startMs",   label: "Start Ms",           wide: false },
    { key: "endMin",    label: "End Min",            wide: false },
    { key: "endSec",    label: "End Sec",            wide: false },
    { key: "endMs",     label: "End Ms",             wide: false },
    { key: "sourceDesc",label: "Source Description", wide: true  } ]

; ------------------------------------------------------------------- state --

bridge      := ExcelBridge()
index       := 1                ; 1-based position in bridge.rows
values      := Map()            ; field key -> string to type
sent        := Map()            ; field key -> true once dispensed
paused      := false
guiHwnd     := 0
lastHwnd    := 0                ; last foreground window that was not the guide
lastCtrl    := 0                ; and the control focused inside it
ui          := Map()            ; control name -> Gui control

; ------------------------------------------------------------------ startup --

Main()

Main() {
    global bridge, index
    try {
        names := bridge.Attach()
        ; An argument names the workbook outright: "CaptionFiller.ahk Green.xlsx"
        chosen := A_Args.Length ? A_Args[1] : ChooseWorkbook(names)
        if (chosen = "")
            ExitApp
        bridge.BindWorkbook(chosen)
        bridge.Load()
    } catch as e {
        MsgBox(e.Message, APP_TITLE, "Iconx")
        ExitApp
    }

    BuildGui()

    ; Resume where we left off, if it is still a valid position.
    remembered := Integer(IniRead(INI, "state", "index", "1"))
    index := (remembered >= 1 && remembered <= bridge.rows.Length) ? remembered : 1

    ShowCaption()
    SetTimer(TrackFocus, 250)
}

/*  Five workbooks open is normal, so always let the user say which one.  */
ChooseWorkbook(names) {
    candidates := []
    for name in names {
        if (SubStr(name, 1, 1) = "~")           ; Excel lock files
            continue
        if (name = "PERSONAL.XLSB")             ; the macro workbook, never data
            continue
        candidates.Push(name)
    }
    if !candidates.Length {
        MsgBox("No workbooks are open in Excel.`n`nOpen your annotation workbook first.", APP_TITLE, "Iconx")
        return ""
    }
    if (candidates.Length = 1)
        return candidates[1]

    preferred := IniRead(INI, "state", "workbook", "")
    picked := ""
    chooser := Gui("+AlwaysOnTop -MinimizeBox", APP_TITLE)
    chooser.SetFont("s10", "Segoe UI")
    chooser.Add("Text",, "Which workbook holds the annotations?")
    list := chooser.Add("ListBox", "w360 r" Min(10, candidates.Length), candidates)
    list.Choose(1)
    for i, name in candidates
        if (name = preferred)
            list.Choose(i)
    ok := chooser.Add("Button", "w110 Default", "Use this")
    cancel := chooser.Add("Button", "x+10 yp w110", "Cancel")
    ok.OnEvent("Click", (*) => (picked := list.Text, chooser.Destroy()))
    cancel.OnEvent("Click", (*) => chooser.Destroy())
    chooser.OnEvent("Close", (*) => chooser.Destroy())
    chooser.Show()
    WinWaitClose("ahk_id " chooser.Hwnd)
    if (picked != "")
        IniWrite(picked, INI, "state", "workbook")
    return picked
}

; --------------------------------------------------------------------- view --

BuildGui() {
    global ui, guiHwnd

    g := Gui("AlwaysOnTop -MinimizeBox -MaximizeBox +E0x08000000 -DPIScale", APP_TITLE)
    g.SetFont("s9", "Segoe UI")
    g.MarginX := 10, g.MarginY := 8
    ui["gui"] := g

    ; --- orientation and navigation ---
    ui["trackText"] := g.Add("Text", "w320 h18", "Track -")
    b := g.Add("Button", "x+6 yp-2 w70 h24", Glyph.PREV " Track")
    b.OnEvent("Click", (*) => GoTrack(-1))
    ui["prevTrack"] := b
    b := g.Add("Button", "x+4 yp w70 h24", "Track " Glyph.NEXT)
    b.OnEvent("Click", (*) => GoTrack(1))
    ui["nextTrack"] := b

    ui["captionText"] := g.Add("Text", "xm y+6 w320 h18", "Caption -")
    b := g.Add("Button", "x+6 yp-2 w70 h24", Glyph.PREV " Cap")
    b.OnEvent("Click", (*) => GoCaption(-1))
    ui["prevCaption"] := b
    b := g.Add("Button", "x+4 yp w70 h24", "Cap " Glyph.NEXT)
    b.OnEvent("Click", (*) => GoCaption(1))
    ui["nextCaption"] := b

    g.Add("Text", "xm y+8 w480 h1 0x10")     ; separator

    ; --- fill buttons ---
    ; Wide fields get their own row; the six time boxes sit three to a row.
    for field in FIELDS {
        key := field.key
        if field.wide {
            btn := g.Add("Button", "xm y+6 w150 h26", "   " field.label)
            ui["value_" key] := g.Add("Edit", "x+8 yp+2 w322 h22 ReadOnly -E0x200")
        } else {
            newRow := (key = "startMin" || key = "endMin")
            opts := newRow ? "xm y+6 w104 h26" : "x+6 yp w104 h26"
            btn := g.Add("Button", opts, "   " field.label)
            ui["value_" key] := g.Add("Edit", "x+4 yp+2 w48 h22 ReadOnly Center -E0x200")
        }
        btn.OnEvent("Click", Fill.Bind(key))
        ui["btn_" key] := btn
    }

    g.Add("Text", "xm y+10 w480 h1 0x10")

    ; --- status and options ---
    ui["status"] := g.Add("Text", "xm y+6 w480 h18", "Ready.")

    ui["replace"] := g.Add("Checkbox", "xm y+6 Checked", "Replace field contents")
    ui["padMs"]   := g.Add("Checkbox", "x+12 yp", "Pad ms to 3")
    ui["follow"]  := g.Add("Checkbox", "x+12 yp Checked", "Follow in Excel")

    b := g.Add("Button", "xm y+8 w90 h26", "Reload")
    b.OnEvent("Click", (*) => Reload_())
    ui["pause"] := g.Add("Button", "x+6 yp w90 h26", "Pause")
    ui["pause"].OnEvent("Click", (*) => TogglePause())

    ; Restore remembered options and position.
    ui["replace"].Value := Integer(IniRead(INI, "options", "replace", "1"))
    ui["padMs"].Value   := Integer(IniRead(INI, "options", "padMs", "0"))
    ui["follow"].Value  := Integer(IniRead(INI, "options", "follow", "1"))

    x := IniRead(INI, "window", "x", "")
    y := IniRead(INI, "window", "y", "")
    g.OnEvent("Close", (*) => Shutdown())
    if (x != "" && y != "")
        g.Show("NoActivate x" x " y" y)
    else
        g.Show("NoActivate")
    guiHwnd := g.Hwnd
}

/*  Loads the current row into the buttons.  */
ShowCaption() {
    global bridge, index, values, sent, ui

    record := bridge.Refresh(index)
    start := SplitTime(record.startTime)
    end   := SplitTime(record.endTime)
    pad   := ui["padMs"].Value

    values := Map()
    values["caption"]    := record.caption
    values["startMin"]   := start.min
    values["startSec"]   := start.sec
    values["startMs"]    := PadMs(start.ms, pad)
    values["endMin"]     := end.min
    values["endSec"]     := end.sec
    values["endMs"]      := PadMs(end.ms, pad)
    values["sourceDesc"] := record.sourceDesc

    sent := Map()

    trackLabel := (record.trackNo = "" && record.trackName = "")
        ? "Track  -  (this sheet has no track information)"
        : Format("Track  {1} / {2}    {3}"
            , TrackPosition(record.trackNo), bridge.tracks.Length, record.trackName)
    ui["trackText"].Value := trackLabel
    ui["captionText"].Value := Format("Caption  {1} / {2}    (sheet row {3})"
        , index, bridge.rows.Length, record.excelRow)

    for field in FIELDS {
        value := values[field.key]
        ui["value_" field.key].Value := value
        ui["btn_" field.key].Enabled := (value != "")
    }

    ui["prevCaption"].Enabled := (index > 1)
    ui["nextCaption"].Enabled := (index < bridge.rows.Length)
    ui["prevTrack"].Enabled := (AdjacentTrack(-1) != 0)
    ui["nextTrack"].Enabled := (AdjacentTrack(1) != 0)

    MarkNext()
    Status(Shorten(record.caption, 70))     ; stale "Sent ..." would misreport this caption
    if ui["follow"].Value
        bridge.ShowRow(index)
    IniWrite(index, INI, "state", "index")
}

PadMs(ms, pad) {
    if (ms = "")
        return ""
    return pad ? Format("{:03}", ms) : String(ms)
}

/*  The leading glyph on each button: sent, next up, or neither.  */
MarkNext() {
    global ui, sent, values
    nextKey := ""
    for field in FIELDS {
        if (values[field.key] != "" && !sent.Has(field.key)) {
            nextKey := field.key
            break
        }
    }
    for field in FIELDS {
        mark := sent.Has(field.key) ? Glyph.SENT " " : (field.key = nextKey ? Glyph.NEXT " " : "   ")
        ui["btn_" field.key].Text := mark field.label
    }
}

; ---------------------------------------------------------------- behaviour --

/*  Keeps track of where a value should go: the last window that was in front
    and was not this guide, plus the control focused inside it.  */
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
    global values, sent, ui, paused, guiHwnd, lastHwnd, lastCtrl

    ; SendText returns before the receiving window has processed the keystrokes.
    ; Without this guard, clicking a second field quickly can land its Ctrl+A
    ; in the middle of the previous value and shred both.
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
    ; Refuse to type into Excel. Otherwise clicking a cell to check something,
    ; then clicking a fill button, would overwrite the annotation data itself.
    try {
        if (WinGetProcessName(lastHwnd) = "EXCEL.EXE") {
            Status("That would type into Excel. Click the website field first.")
            return
        }
    }

    ; The safety net. With WS_EX_NOACTIVATE the guide never took focus and this
    ; does nothing; without it, this puts focus back where the value belongs.
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
        SendText(value)      ; SendText, never Send: captions contain { } ! ^ + #
        ; Let the target drain its message queue before another fill is allowed.
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

/*  First row of the next/previous track number, or 0 if there is none.
    Uses the sorted list of distinct track numbers, so gaps do not matter.  */
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
            wanted := trackNo          ; keep the largest that is still smaller
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

Reload_() {
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
    for field in FIELDS
        if (field.key = key)
            return field.label
    return key
}

Shorten(text, limit) {
    text := StrReplace(StrReplace(text, "`r", " "), "`n", " ")
    return StrLen(text) > limit ? SubStr(text, 1, limit - 1) Glyph.ELL : text
}

Shutdown() {
    global ui, index
    try {
        ui["gui"].GetPos(&x, &y)
        IniWrite(x, INI, "window", "x")
        IniWrite(y, INI, "window", "y")
        IniWrite(ui["replace"].Value, INI, "options", "replace")
        IniWrite(ui["padMs"].Value, INI, "options", "padMs")
        IniWrite(ui["follow"].Value, INI, "options", "follow")
        IniWrite(index, INI, "state", "index")
    }
    ExitApp
}
