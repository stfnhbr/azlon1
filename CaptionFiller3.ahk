#Requires AutoHotkey v2.0
#SingleInstance Force
/*
    AutoNyx Caption Filler III
    --------------------------
    A floating guide for the "Soup EE" column schema, same sheets and same
    Excel bridge as Caption Filler II, laid out differently.

    What is different from II:

    1. The columns sit in four blocks between the two dividers, grouped by what
       they are rather than by where they fall in the sheet:

           A   Caption Number, TrackNumber          read-only context
           B   Track, TrackDescription              the track, filled
           C   AnnotationText, SourceVisibility,    the annotation, filled
               SourceDescription, Prominence          and its context
           D   StartTime(s), EndTime(s)             the times, filled

    2. No Min buttons, and no decimal seconds either. A time offers Sec and Ms
       and nothing else: its column name is a tag, not a button, because 90.500
       is not a shape any box on the website takes.

       Sec therefore holds the TOTAL whole seconds - 90.5 s types as Sec 90,
       Ms 500, not as 30 and 500. Filler II split at 60 because a Min box took
       the overflow; with no Min box there is nowhere for it to go.

    3. Each block sits on a tint of its own, so the eye finds a block before it
       reads a label. The tint is a band behind the whole block, added last and
       pushed to the back, with every control painting on top of it.

    4. SourceVisibility and Prominence are not typed anywhere, so instead of a
       value box they show the value itself as a colored chip, cut to the width
       of its own text: sharp corners for visibility, rounded for prominence.
       A value the sheet uses but this file has no color for still shows in
       full, in plain black on the block's tint.

    How you use it:
        1. click the target box on the website  (the browser keeps focus)
        2. click the matching button here       (the value is typed into it)

    The guide never clicks anything in the browser, never presses Tab, and
    never submits the form. It only types into the field you focused.

    This is a SEPARATE tool from CaptionFiller.ahk and CaptionFiller2.ahk. It
    shares lib\ExcelBridge2.ahk with the second, read-only, and shares no code
    with either script, on purpose: those two are trusted working software and
    a change made here must not be able to reach them.
*/

#Include lib\ExcelBridge2.ahk

; ---------------------------------------------------------------- constants --

APP_TITLE := "AutoNyx Caption Filler III"
INI       := A_ScriptDir "\CaptionFiller3.ini"

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

/*  Row widths. The first column is 170 throughout - tags and fill buttons all
    line up - and every row spans the same 578px as the dividers above and
    below it:

        fill / read     170 + 8 + 400                              = 578
        time            170 + 8 + 120 + 4 + 71 + 10 + 120 + 4 + 71 = 578

    Everything the GUI reads has to be assigned before Main() runs, which is why
    this lives up here with the rest of the constants rather than beside
    BuildGui: the auto-execute section stops at the Main() call below, and a
    global assigned after it is still empty when the window is built.  */
class Width {
    static LABEL := 170       ; first column: a tag or the row's own button
    static VALUE := 400       ; the cell as it stands
    static PART  := 120       ; a Sec or Ms button
    static PVAL  :=  71       ; and the box beside it
    static ROW   := 578       ; dividers, status line

    /*  Heights. BOX_H is why the value boxes are not as tall as the controls
        beside them: a single-line Edit draws its text at the top of its box
        and leaves the rest empty, whatever height you give it, and there is no
        vertical-centre flag for it. At 22 the ink sat 3px from the top and 7px
        from the bottom, which is what reads as top-aligned. So the box is cut
        to the line instead, and then dropped to sit centred on its taller
        neighbour.

        18 measured, not guessed. In Segoe UI 9pt the ink of ordinary text runs
        from 3px below the top of the box to 14px below it - capitals start at
        3, descenders end at 14 - so a box of 18 leaves 3 above and 3 below.
        Text carrying an umlaut or an accent reaches up to 1px instead of 3 and
        so sits a pixel high, which is invisible; nothing clips either way.  */
    static BOX_H :=  18       ; every value box
    static BTN_H :=  26       ; rows led by a button
    static LBL_H :=  22       ; rows led by a label
}

/*  Vertical space before a row: a little inside a block, more between two.
    Lead, not Gap: the row loop keeps a "gap" variable, and a class name
    collides with any same-named variable because identifiers here are
    case-insensitive - the same reason Sym above is not called G.  */
class Lead {
    static INSIDE := 5
    static BLOCK  := 16
}

/*  The band painted behind a block, a little wider and taller than the rows
    standing on it.  */
class Band {
    static X   :=   6                  ; a touch left of the 10px margin
    static W   := 590                  ; and the same touch past the right edge
    static PAD :=   4                  ; above the first row, below the last
}

/*  Every color in the window, as a six-digit hex string without the "#".

    BLOCK tints a whole block. SOURCE_VIS and PROMINENCE color one chip each,
    keyed by the cell's exact text - a value not listed here is not an error,
    it simply shows uncolored.

    Prominent was given as "#ff699ff", which is seven digits and cannot be a
    color. Moderate and Slight are ffcc99 and ffff99, so the three are a
    red-orange-yellow ramp of the form ff<XX>99, and ff9999 is the reading that
    completes it. If the intent was ff6699, this one line is the whole change.  */
class Tint {
    static PAPER := "FFFFFF"    ; the boxes a value is still sent from

    static BLOCK := Map(
        "A", "E0F2FE",      ; Caption Number, TrackNumber
        "B", "D1FAE5",      ; Track, TrackDescription
        "C", "FEF9C3",      ; AnnotationText and its context
        "D", "FFE4E6")      ; StartTime(s), EndTime(s)

    static SOURCE_VIS := Map(
        "Visible source"                , "006600",
        "Intermittently on screen"      , "663300",
        "Not visible / offscreen source", "9A0000",
        "Unclear"                       , "515151")

    /*  Slight is deliberately lighter than block C's own FEF9C3, so the least
        prominent value reads as the faintest thing in the block rather than as
        a chip shouting the same yellow as the band behind it.  */
    static PROMINENCE := Map(
        "Prominent", "FF9999",
        "Moderate" , "FFCC99",
        "Slight"   , "FFFFE8")
}

/*  A chip: the value on a colored ground, cut minimally larger than its text.
    ROUND is the corner diameter fed to CreateRoundRectRgn, 0 for sharp.  */
class Chip {
    static PADX  := 8
    static PADY  := 3
    static SHARP := 0
    static ROUND := 10
    static INK_ON_DARK  := "FFFFFF"     ; SourceVisibility, on its dark grounds
    static INK_ON_LIGHT := "000000"     ; Prominence, on its pale ones
}

/*  Every column, in screen order. field must match ExcelBridge2.Fields.
    kind decides the row:
        "read"  shown only, no button
        "chip"  shown only, as a colored chip - palette, ink and corner apply
        "fill"  one button typing the cell as it stands
        "time"  a tag naming the column, then one button per time part
    block groups the rows on screen and picks the tint behind them.
    font, where a row has one, restyles that row's value box: weight, slant and
    colour, over whatever typeface the window is already using.
    prefix builds the part keys ("start" + "Sec"), short names them on screen.  */
FIELDS := [
    { field: "captionNo" , label: "Caption Number"    , kind: "read", block: "A" },
    { field: "trackNo"   , label: "TrackNumber"       , kind: "read", block: "A" },

    { field: "track"     , label: "Track"             , kind: "fill", block: "B" },
    { field: "trackDesc" , label: "TrackDescription"  , kind: "fill", block: "B" },

    { field: "annotation", label: "AnnotationText"    , kind: "fill", block: "C"
                         , font: "Bold Italic c0000CC" },
    { field: "sourceVis" , label: "SourceVisibility"  , kind: "chip", block: "C"
                         , palette: Tint.SOURCE_VIS
                         , ink: Chip.INK_ON_DARK , corner: Chip.SHARP },
    { field: "sourceDesc", label: "SourceDescription" , kind: "fill", block: "C" },
    { field: "prominence", label: "Prominence"        , kind: "chip", block: "C"
                         , palette: Tint.PROMINENCE
                         , ink: Chip.INK_ON_LIGHT, corner: Chip.ROUND },

    { field: "startTime" , label: "StartTime(s)"      , kind: "time", block: "D"
                         , prefix: "start", short: "Start" },
    { field: "endTime"   , label: "EndTime(s)"        , kind: "time", block: "D"
                         , prefix: "end"  , short: "End"   } ]

/*  The two boxes a time is typed into, in website order. No Min: see the note
    at the top, and TimeParts below, which adds the minutes back into Sec.  */
TIME_PARTS := [
    { suffix: "Sec", label: "Sec" },
    { suffix: "Ms" , label: "Ms"  } ]

/*  One entry per button, in click order, derived from the tables above so the
    two cannot drift apart. key indexes values and sent, field names the column
    the value comes from, and hint marks the buttons the Sym.NEXT pointer walks.

    caption is what fits on the button, label is what the status line says: the
    time buttons read "Sec" in a row already introduced by StartTime(s), but
    "Sent Sec: 90" downstairs would not say which time it came from.  */
BUTTONS := FillButtons()

FillButtons() {
    global FIELDS, TIME_PARTS
    list := []
    for entry in FIELDS {
        if (entry.kind = "fill")
            list.Push({ key: entry.field, field: entry.field
                      , caption: entry.label, label: entry.label, hint: true })
        else if (entry.kind = "time") {
            /*  No button for the decimal seconds. The website takes a time in
                boxes, never as 90.500, so a button offering it could only ever
                type the wrong shape into the right field. The column name is
                left as a tag over the parts that are actually sent.  */
            for part in TIME_PARTS
                list.Push({ key: entry.prefix part.suffix, field: entry.field
                          , caption: part.label
                          , label: entry.short " " part.label, hint: true })
        }
    }
    return list
}

; ------------------------------------------------------------------- state --

bridge   := ExcelBridge2()
index    := 1                ; 1-based position in bridge.rows
values   := Map()            ; field or button key -> string currently shown
sent     := Map()            ; button key -> true once dispensed
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

    ; An argument names the workbook outright: CaptionFiller3.ahk "Soup EE.xlsx"
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

    win.Add("Text", "xm y+8 w" Width.ROW " h1 0x10")

    /*  --- the ten columns, in four blocks ---

        first is the row's opening control, which is also its tallest: a block's
        band is measured from those, not from the value boxes beside them.

        Note what none of these locals are called. Identifiers here are
        case-insensitive, so a variable named tint, lead, chip or band would
        shadow the class of that name for the rest of the function and every
        reading of it would fail. The classes keep the nouns; the locals take
        shade, first, badge and strip. Same trap the Sym comment describes.  */
    ; How far a value box drops to sit centred on the taller control beside it.
    onButton := (Width.BTN_H - Width.BOX_H) // 2
    onLabel  := (Width.LBL_H - Width.BOX_H) // 2

    lastBlock := ""
    bounds := Map()
    for entry in FIELDS {
        key := entry.field
        shade := Tint.BLOCK[entry.block]
        gap := (lastBlock = "" || entry.block = lastBlock) ? Lead.INSIDE : Lead.BLOCK
        lastBlock := entry.block

        if (entry.kind = "fill") {
            first := win.Add("Button", "xm y+" gap " w" Width.LABEL " h26", "   " entry.label)
            first.OnEvent("Click", Fill.Bind(key))
            ui["btn_" key] := first
            ; White on purpose, and said outright: a ReadOnly Edit left to
            ; itself comes up in the dialog grey, which on a tinted band reads
            ; as a hole rather than as a field. White keeps every box you can
            ; still send from legible against whatever colour it stands on.
            ui["value_" key] := win.Add("Edit", "x+8 yp+" onButton " w" Width.VALUE " h" Width.BOX_H " ReadOnly -E0x200 +Background" Tint.PAPER)
        } else if (entry.kind = "time") {
            /*  A tag, not a button. The decimal seconds it used to offer are
                not a shape any box on the website takes, so there is nothing
                here to send - only the Sec and Ms parts below it are.

                The tag is as tall as the buttons beside it, unlike the tags on
                the read and chip rows, so that the band measured from it still
                covers them.  */
            first := win.Add("Text", "xm y+" gap " w" Width.LABEL " h" Width.BTN_H " Center +0x200 +Background" shade
                           , entry.label)
            for part in TIME_PARTS {
                partKey := entry.prefix part.suffix
                ; The first button follows the tag; the later ones have to climb
                ; back out of the value box beside them onto the button line.
                step := (A_Index = 1) ? "x+8 yp" : "x+10 yp-" onButton
                btn := win.Add("Button", step " w" Width.PART " h" Width.BTN_H, "   " part.label)
                btn.OnEvent("Click", Fill.Bind(partKey))
                ui["btn_" partKey] := btn
                ui["value_" partKey] := win.Add("Edit", "x+4 yp+" onButton " w" Width.PVAL " h" Width.BOX_H " ReadOnly Center -E0x200 +Background" Tint.PAPER)
            }
        } else if (entry.kind = "chip") {
            ; Nothing is ever typed from these, so there is no box to type from.
            ; The chip is built full width and cut down to its text by PaintChip,
            ; which runs again on every caption because the text changes.
            first := win.Add("Text", "xm y+" gap " w" Width.LABEL " h" Width.LBL_H " Center +0x200 +Background" shade
                           , entry.label)
            badge := win.Add("Text", "x+8 yp w" Width.VALUE " h" Width.LBL_H " +0x200 Center")
            badge.GetPos(&badgeX, &badgeY, &badgeW, &badgeH)
            ui["value_" key] := badge
            ; Where the chip lives no matter what size it is cut to.
            ui["slot_" key] := { x: badgeX, y: badgeY, h: badgeH }
        } else {
            ; Read-only context: shown, but nothing to click. Tinted to match
            ; the band so it reads as part of the block rather than as a field.
            first := win.Add("Text", "xm y+" gap " w" Width.LABEL " h" Width.LBL_H " Center +0x200 +Background" shade
                           , entry.label)
            ui["value_" key] := win.Add("Edit", "x+8 yp+" onLabel " w" Width.VALUE " h" Width.BOX_H " ReadOnly -E0x200 +Background" shade)
        }

        /*  A row may restyle its own value box - see the font entry in FIELDS.

            Options only, no typeface, so the box keeps Segoe UI and only its
            weight, slant and colour change. This is the half that does the
            work: the font property on its own is inert, and dropping this line
            leaves AnnotationText looking like every other row.  */
        if entry.HasOwnProp("font")
            ui["value_" key].SetFont(entry.font)

        first.GetPos(&firstX, &firstY, &firstW, &firstH)
        if bounds.Has(entry.block)
            bounds[entry.block].bottom := firstY + firstH
        else
            bounds[entry.block] := { top: firstY, bottom: firstY + firstH }
    }

    win.Add("Text", "xm y+10 w" Width.ROW " h1 0x10")

    ; --- status and options ---
    ui["status"] := win.Add("Text", "xm y+6 w" Width.ROW " h18", "Ready.")

    ui["replace"] := win.Add("Checkbox", "xm y+6 Checked", "Replace field contents")
    ui["padMs"]   := win.Add("Checkbox", "x+14 yp", "Pad ms to 3")
    ui["follow"]  := win.Add("Checkbox", "x+14 yp Checked", "Follow in Excel")
    ui["padMs"].OnEvent("Click", (*) => RepadMs())

    btn := win.Add("Button", "xm y+8 w96 h26", "Reload")
    btn.OnEvent("Click", (*) => ReloadSheet())
    ui["pause"] := win.Add("Button", "x+6 yp w96 h26", "Pause")
    ui["pause"].OnEvent("Click", (*) => TogglePause())

    ui["replace"].Value := Integer(IniRead(INI, "options", "replace", "1"))
    ui["padMs"].Value   := Integer(IniRead(INI, "options", "padMs", "0"))
    ui["follow"].Value  := Integer(IniRead(INI, "options", "follow", "1"))

    /*  The bands behind the blocks, last of all.

        Last, because a control added here takes its position from the one
        before it, and a band dropped in mid-layout would push every row that
        follows out of place. Adding them at the end costs nothing but z-order,
        and z-order is then set outright: SendToBack puts each band underneath
        every control already on the window.

        Being underneath is not by itself enough. A control with no
        WS_CLIPSIBLINGS paints its whole rectangle whatever is stacked above it,
        so a band at the very bottom of the pile still wipes out every row
        standing on it - the window comes up as four colored slabs and nothing
        else. 0x4000000 is that style: it clips the band to what is not already
        covered by a sibling above it, which is exactly the margins and the gaps
        between the rows.  */
    for letter, hex in Tint.BLOCK {
        if !bounds.Has(letter)
            continue
        edge := bounds[letter]
        strip := win.Add("Text", "x0 y0 w10 h10 +0x4000000 +Background" hex)
        strip.Move(Band.X, edge.top - Band.PAD, Band.W, edge.bottom - edge.top + 2 * Band.PAD)
        SendToBack(strip.Hwnd)
        ui["band_" letter] := strip
    }

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
    for entry in FIELDS {
        values[entry.field] := record.%entry.field%
        if (entry.kind = "time") {
            parts := TimeParts(record.%entry.field%, ui["padMs"].Value)
            for part in TIME_PARTS
                values[entry.prefix part.suffix] := parts[part.suffix]
        }
    }

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
        shown := present ? value : "(not in this sheet)"
        if (entry.kind = "chip") {
            PaintChip(entry, shown)
            continue
        }
        /*  A time row owns no control of its own any more - its tag is fixed
            text and the decimal seconds are gone - so it is only its parts that
            get filled in here. Blank when the cell holds something that is not
            a time at all, which leaves the buttons disabled rather than
            offering to type nonsense.  */
        if (entry.kind = "time") {
            for part in TIME_PARTS {
                partKey := entry.prefix part.suffix
                ui["value_" partKey].Value := present ? values[partKey] : ""
                ui["btn_" partKey].Enabled := (present && values[partKey] != "")
            }
            continue
        }
        ui["value_" entry.field].Value := shown
        if (entry.kind = "read")
            continue
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

/*  Draws one value as a chip: colored ground, cut to its own text plus a little
    padding, sitting on the row's centre line.

    A value the palette does not list - a blank cell, "(not in this sheet)", or
    a wording nobody has met yet - is not an error and is not hidden. It shows
    in black on the block's own tint, which is to say as plain text on the band,
    and the reader simply sees no chip there.

    Re-run for every caption, because the text changes and the chip is cut to
    the text. The band underneath is redrawn first: shrinking a chip uncovers
    the band, and the uncovered strip has to be repainted by something.  */
PaintChip(entry, text) {
    global ui

    ; badge, not chip: a local named chip would shadow the Chip class below.
    badge := ui["value_" entry.field]
    slot  := ui["slot_" entry.field]
    shade := Tint.BLOCK[entry.block]

    known  := (text != "" && entry.palette.Has(text))
    ground := known ? entry.palette[text] : shade
    ink    := known ? entry.ink : Chip.INK_ON_LIGHT

    size := MeasureText(badge, text)
    w := size.w + 2 * Chip.PADX
    h := size.h + 2 * Chip.PADY

    badge.Value := text
    badge.Opt("+Background" ground " c" ink)
    ; The region is in the control's own coordinates, so it has to be cut again
    ; every time the control is resized. Sharp corners get no region at all.
    badge.Move(slot.x, slot.y + (slot.h - h) // 2, w, h)
    SetCorners(badge.Hwnd, w, h, entry.corner)

    if ui.Has("band_" entry.block)
        ui["band_" entry.block].Redraw()
    badge.Redraw()
}

/*  Width and height of text as the control's own font would draw it. Asking
    the control for its font matters: measuring with the default system font
    would cut every chip to the wrong size.  */
MeasureText(ctrl, text) {
    static WM_GETFONT := 0x31

    if (text = "")
        return { w: 0, h: 0 }

    font := DllCall("SendMessage", "Ptr", ctrl.Hwnd, "UInt", WM_GETFONT
                  , "Ptr", 0, "Ptr", 0, "Ptr")
    dc := DllCall("GetDC", "Ptr", ctrl.Hwnd, "Ptr")
    previous := font ? DllCall("SelectObject", "Ptr", dc, "Ptr", font, "Ptr") : 0

    size := Buffer(8, 0)
    DllCall("gdi32\GetTextExtentPoint32W", "Ptr", dc, "WStr", text
          , "Int", StrLen(text), "Ptr", size)

    if previous
        DllCall("SelectObject", "Ptr", dc, "Ptr", previous)
    DllCall("ReleaseDC", "Ptr", ctrl.Hwnd, "Ptr", dc)

    return { w: NumGet(size, 0, "Int"), h: NumGet(size, 4, "Int") }
}

/*  Rounds a control's corners by clipping it to a rounded rectangle, or clears
    any previous clipping when the diameter is 0. The window owns the region
    once it is set, so it must not be deleted here.  */
SetCorners(hwnd, w, h, diameter) {
    if !diameter {
        DllCall("SetWindowRgn", "Ptr", hwnd, "Ptr", 0, "Int", true)
        return
    }
    region := DllCall("gdi32\CreateRoundRectRgn", "Int", 0, "Int", 0
                    , "Int", w + 1, "Int", h + 1
                    , "Int", diameter, "Int", diameter, "Ptr")
    if region
        DllCall("SetWindowRgn", "Ptr", hwnd, "Ptr", region, "Int", true)
}

/*  Puts a control underneath every other control on the window.  */
SendToBack(hwnd) {
    static HWND_BOTTOM := 1
    static SWP_NOSIZE := 0x1, SWP_NOMOVE := 0x2, SWP_NOACTIVATE := 0x10
    DllCall("SetWindowPos", "Ptr", hwnd, "Ptr", HWND_BOTTOM
          , "Int", 0, "Int", 0, "Int", 0, "Int", 0
          , "UInt", SWP_NOSIZE | SWP_NOMOVE | SWP_NOACTIVATE)
}

/*  The two boxes a time is typed into, keyed by the part suffix.

    SplitTime2 hands back minutes separately, the way Filler II's form wanted
    them; with no Min box here the minutes are folded back into the seconds, so
    Sec is the total: 90.5 s gives Sec 90, not 30. Leaving them out instead
    would type a time an hour short of the truth and look perfectly reasonable
    doing it.

    pad writes the milliseconds three digits wide: 48 becomes 048.  */
TimeParts(seconds, pad) {
    parts := SplitTime2(seconds)
    if (parts.sec = "")                 ; not a time at all
        return Map("Sec", "", "Ms", "")
    return Map(
        "Sec", String(parts.min * 60 + parts.sec),
        "Ms" , pad ? Format("{:03}", parts.ms) : String(parts.ms))
}

/*  Re-renders only the millisecond boxes when the pad is toggled. A full
    ShowCaption would do it too, but it would also clear the check marks of
    everything already typed into this caption.  */
RepadMs() {
    global ui, values

    pad := ui["padMs"].Value
    for entry in FIELDS {
        if (entry.kind != "time")
            continue
        key := entry.prefix "Ms"
        if (!values.Has(key) || values[key] = "")
            continue
        ms := Integer(values[key])
        values[key] := pad ? Format("{:03}", ms) : String(ms)
        ui["value_" key].Value := values[key]
    }
}

/*  Leading glyph per button: sent, next up, or neither.  */
MarkNext() {
    global ui, sent, values

    nextKey := ""
    for button in BUTTONS {
        if (button.hint && values[button.key] != "" && !sent.Has(button.key)) {
            nextKey := button.key
            break
        }
    }
    for button in BUTTONS {
        mark := sent.Has(button.key) ? Sym.SENT " "
              : (button.key = nextKey ? Sym.NEXT " " : "   ")
        ui["btn_" button.key].Text := mark button.caption
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
    for button in BUTTONS
        if (button.key = key)
            return button.label
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
        IniWrite(ui["padMs"].Value, INI, "options", "padMs")
        IniWrite(ui["follow"].Value, INI, "options", "follow")
        IniWrite(index, INI, "state", "index")
    }
    ExitApp
}
