# AutoNyx — annotation tools

Turning Audacity label annotations into spreadsheets. Two ways in.

## 1. Hotkey, straight out of a live project (preferred)

Press **Ctrl+Shift+Alt+K** while Audacity is focused. A save dialog opens with the
filename already set to the project name; confirm it and the workbook opens in
Excel.

All label tracks are merged and sorted by start time. See [Columns](#columns).

To arm the hotkey, double-click **`Audacity annotation hotkey.ahk`** (an "H"
icon appears in the tray). To have it armed after every login, press `Win+R`,
run `shell:startup`, and drop a shortcut to that file in the folder that opens.

This route reads Audacity directly through `mod-script-pipe`, so there is no
Export Labels step — and unlike Audacity's own `.txt` export, it keeps the name
of the label track each annotation came from.

Requires `mod-script-pipe` set to **Enabled** under
*Edit > Preferences > Modules* (already the case here), and a restart of
Audacity after changing that setting.

## 2. Drag and drop an exported .txt

Drop one or more Audacity label `.txt` files onto **`Drop labels here.cmd`**.
Each produces a sibling `.xlsx` and opens it.

Use this for files exported earlier or received from someone else. Audacity's
`.txt` export flattens every label track into one list, so Track Number and
Track Name come out empty on this route — the layout is otherwise identical, so
later steps see one consistent shape either way.

## Columns

Both routes write the same ten columns, with autofilter and a frozen header row.

| # | Column | Filled by |
|---|---|---|
| 1 | Caption Number | generated — 1..n down the sheet |
| 2 | Track Number | Audacity — the label track's position (hotkey route only) |
| 3 | Track Name | Audacity — e.g. `3 Bird Sounds` (hotkey route only) |
| 4 | Track Description | *empty, for you* |
| 5 | Caption | Audacity — the label text |
| 6 | Start Time | Audacity |
| 7 | End Time | Audacity |
| 8 | Source Visibility | *empty, for you* |
| 9 | Source Description | *empty, for you* |
| 10 | Prominence | *empty, for you* |

To add, rename or reorder columns, edit the single `$AnnotationColumns` table at
the top of `lib\AnnotationWorkbook.ps1` — headers, number formats and widths all
come from there, and both routes follow it.

## 3. Numbered captions in, one label track per category

Press **Ctrl+Shift+Alt+I** while Audacity is focused. Pick a caption file and it
becomes one label track per category number.

The annotation site hands out files like `Hut R0_sound_nouns.txt`, tab separated,
where a number before the dash says which category the caption belongs to:

```text
0.000	30.000	8 – Insect chirping
0.200	0.933	1 – Male speech
1.767	4.767	2 – Male speech
```

Audacity's own *File > Import > Labels* flattens all of that into a **single**
track, losing the grouping. This builds a track per number instead — named `1`,
`2`, `3`, … in ascending order beneath the audio — and strips the `N – ` prefix
from each caption, since the track already carries the number.

The picker opens in the current project's folder with the matching file already
selected where there is one; the download's `%20` and casing don't have to match
(`HUt%20R0_sound_nouns.txt` is found for a project called `Hut`).

**Two guards, both worth knowing about.** Audacity's scripting pipe always talks
to whichever project is *frontmost*, so with several projects open it is easy to
fire a file at the wrong one:

- if the file's name doesn't match the open project, it asks before importing —
  `Hut R0_sound_nouns.txt` into a project called `Mud` needs confirming;
- if you switch projects between picking the file and the import starting, it
  stops and imports nothing rather than writing into the project that happens to
  be in front.

If the project already has label tracks, it asks once whether to replace them
all or leave the project alone. Replacing is a single undo step.

Afterwards it reads every label back and compares against the file, so a
miscount or a caption on the wrong track is reported rather than left to be
found later in Excel.

The result is indistinguishable from label tracks made by hand, so
Ctrl+Shift+Alt+K exports them exactly as in [Columns](#columns) — Track Name
being `1`, `2`, `3` and so on.

Overlapping captions inside one category are kept as they are; Audacity label
tracks support them and they occur in real files.

## 4. Filling the annotation website — Caption Filler

Double-click **`CaptionFiller.ahk`**. Pick the workbook, and a small always-on-top
guide appears holding one caption at a time.

Two clicks per field, no keyboard, no clipboard:

1. click the target box on the website — the browser keeps focus
2. click the matching button on the guide — the value is typed into that box

The guide never clicks anything in the browser, never presses Tab, and never
submits the form. It only types into the field *you* focused.

**Navigation** — `Cap ◀ / ▶` steps through captions in time order across all
tracks; `Track ◀ / ▶` jumps to the first caption of the previous/next track.
Buttons disable at the ends rather than wrapping.

**Eight fill buttons** — Caption, Start Min/Sec/Ms, End Min/Sec/Ms, Source
Description. A `▶` marks the next un-sent field and a `✓` marks ones already
sent; both reset when you navigate. A button is greyed out when its cell is
empty, so Source Description can't fire on a caption that has none.

Times are split from the sheet's decimal seconds, rounded to milliseconds first
so `59.9996 s` becomes `1 min 0 s 0 ms` rather than an impossible `60` in the
seconds box. Tick **Pad ms to 3** if the site wants `060` instead of `60`.

Options: **Replace field contents** (sends Ctrl+A first, so re-filling a field
overwrites rather than appends), **Follow in Excel** (scrolls the sheet to the
caption you're on), **Pause** (stops all typing), **Reload** (re-reads the sheet
after you edit it). Window position, options and your place in the sheet are
remembered in `CaptionFiller.ini`.

Launch it bound to a specific workbook to skip the picker:

```text
CaptionFiller.ahk Green.xlsx
```

## 5. Caption Filler II — the `Soup EE` column schema

A **second, independent** filler for workbooks using this schema:

```text
Caption Number | TrackNumber | Track | TrackDescription | StartTime(s) |
EndTime(s) | SourceVisibility | SourceDescription | Prominence | AnnotationText
```

Double-click **`CaptionFiller2.ahk`**. It works exactly like the filler above —
click the website field, click the button, the value is typed in — but shows
**all ten columns** and offers fill buttons on **six** of them:

| Fillable | Display only |
|---|---|
| Track | Caption Number |
| TrackDescription | TrackNumber |
| StartTime(s) — plus Min / Sec / Ms | SourceVisibility |
| EndTime(s) — plus Min / Sec / Ms | Prominence |
| SourceDescription | |
| AnnotationText | |

**Times, split.** This schema stores a time as one cell of decimal seconds while
the website takes it in three boxes, so each time row carries `Min`, `Sec` and
`Ms` buttons beside the whole-seconds one — twelve buttons in all. The split
rounds to milliseconds first, so `59.9996 s` becomes `1 min 0 s 0 ms` rather than
an impossible `60` in the seconds box, and minutes keep counting past an hour
(`3661.048 s` → `61 min 1 s 48 ms`). Tick **Pad ms to 3** if the site wants `048`
instead of `48`; toggling it re-writes the ms boxes without clearing the `✓`
marks. A cell holding something that isn't a number leaves all three parts blank
and disabled — the seconds box still shows the text verbatim, so you see what is
actually in the sheet.

The whole-seconds button stays for forms that take a single box, typed to three
decimals (`22.700`, `0.000`) to match the sheet and avoid Excel's floating-point
noise. The `▶` next-up pointer skips it and walks `Min`, `Sec`, `Ms` instead, so
following the pointer can't put `22.700` into a minutes box.

`Cap ◀ / ▶` walks the sheet in row order; `Track ◀ / ▶` jumps to the next or
previous TrackNumber. Neither of these sheets is sorted by track, so the track
buttons work from the sorted list of distinct track numbers rather than from
row adjacency.

**Sheet picking.** A workbook can hold several sheets with this schema — e.g.
`Completion` and `Refinement` — alongside sheets that don't (`Summary`). The
tool scans every open workbook, lists only the sheets that carry all ten
columns, binds silently when exactly one matches, and otherwise asks.

**Repeated header blocks.** Some sheets repeat all ten headers two or three
times across the columns: the data, a variant, then a block of TRUE/FALSE
change-flags. The tool anchors on the *first* `Caption Number` and reads only up
to where the next one begins — so it always takes the leftmost, real block. This
matters: binding to the wrong block would type `FALSE` into the website.

This tool shares no code with `CaptionFiller.ahk`. Each has its own bridge, so a
change made for one schema cannot break the other.

## Files

| File | Role |
|---|---|
| `Audacity annotation hotkey.ahk` | Binds Ctrl+Shift+Alt+K and Ctrl+Shift+Alt+I inside Audacity (AutoHotkey v2) |
| `ExportAnnotations.ps1` | Live export: reads the running project, prompts, writes .xlsx |
| `ImportSoundNouns.ps1` | Live import: a caption `.txt` in, one label track per category number |
| `lib\SoundNouns.ps1` | Parses those caption files and groups them by number |
| `Drop labels here.cmd` | Drag-and-drop target for `.txt` files |
| `AudacityLabelsToExcel.ps1` | Converts label `.txt` files to `.xlsx` |
| `lib\AudacityPipe.ps1` | Named-pipe client for Audacity scripting |
| `lib\AnnotationWorkbook.ps1` | Shared Excel writing (both routes use this) |
| `CaptionFiller.ahk` | Floating guide that types captions into the website |
| `lib\ExcelBridge.ahk` | Reads the open workbook; column lookup by header name |
| `CaptionFiller2.ahk` | Second guide, for the `Soup EE` schema (independent) |
| `lib\ExcelBridge2.ahk` | Its own bridge; leftmost-header-block binding |

## Behaviour worth knowing

- Start/End are written as **numbers**, formatted `0.000` but stored at full
  precision — so they sort and subtract correctly regardless of locale.
- The annotation column is forced to text, so a label starting with `=` or `'`
  is not eaten as a formula. Tabs inside a label survive.
- Frequency-range lines (`\ 8000 12000`) and blank lines in `.txt` exports are
  skipped.
- The drag-and-drop route never overwrites: an existing `Green.xlsx` means the
  new one becomes `Green (2).xlsx`. The hotkey route respects whatever the save
  dialog confirmed.
- A bad file in a dropped batch is reported and skipped; the rest still convert.
- The finished workbook is opened in Excel and brought to the front, not merely
  written to disk. Both routes reuse the Excel instance that wrote the file
  rather than shelling out to the file association, because the hotkey starts
  PowerShell hidden and a shelled-out child can inherit that hidden state.

## Requirements

Excel (uses COM to write real `.xlsx`, not CSV), AutoHotkey v2, Audacity 3.x
with `mod-script-pipe`. All present on this machine.

Note: Audacity 4 is also installed here, but it has no scripting pipe — the
hotkey route works with **Audacity 3.7.8** only.

## Changing the hotkeys

Edit the `^+!k::` (export) or `^+!i::` (import) line in
`Audacity annotation hotkey.ahk`.
`^` = Ctrl, `+` = Shift, `!` = Alt.
