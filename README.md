# AutoNyx — annotation tools

Audacity label annotations to spreadsheets and back again. Two ways in, and — via
[Ctrl+Shift+Alt+N](#3b-workbook-in-label-tracks-out--one-press) — one press back
out.

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
`.txt` export flattens every label track into one list, so TrackNumber and
Track come out empty on this route — the layout is otherwise identical, so
later steps see one consistent shape either way.

## Columns

Both routes write the same ten columns, with autofilter and a frozen header row.
The headers and their order are the `Soup EE` ones, so an export can be opened
straight in [Caption Filler II](#5-caption-filler-ii--the-soup-ee-column-schema)
or III without renaming anything.

| # | Column | Filled by |
|---|---|---|
| 1 | Caption Number | generated — 1..n down the sheet |
| 2 | TrackNumber | Audacity — the label track's position (hotkey route only) |
| 3 | Track | Audacity — e.g. `3 Bird Sounds`, whatever the track is called (hotkey route only). Tracks built by [the import](#3-numbered-captions-in-one-label-track-per-category) are named `<number> <track name>`, so a name put in there comes back out here |
| 4 | TrackDescription | *empty, for you* |
| 5 | StartTime(s) | Audacity |
| 6 | EndTime(s) | Audacity |
| 7 | SourceVisibility | *empty, for you* |
| 8 | SourceDescription | *empty, for you* |
| 9 | Prominence | *empty, for you* |
| 10 | AnnotationText | Audacity — the label text |

To add, rename or reorder columns, edit the `$AnnotationLayouts` table at the top
of `lib\AnnotationWorkbook.ps1` — headers, number formats and widths all come
from there, and both routes follow it.

### Going back to the old headers

The layout above replaced an earlier one — `Caption Number | Track Number |
Track Name | Track Description | Caption | Start Time | End Time | Source
Visibility | Source Description | Prominence` — which is the set
[Caption Filler](#4-filling-the-annotation-website--caption-filler) reads.
Both are kept, so switching is one word: in `lib\AnnotationWorkbook.ps1`, set

```powershell
$script:AnnotationLayout = 'Legacy'      # 'SoupEE' is the current default
```

Both routes follow that line. For a single export without changing it, run
`ExportAnnotations.ps1 -Layout Legacy`. The pre-change file is also kept
verbatim as `lib\AnnotationWorkbook.ps1.bak-original`, and in git history.

## 3. Numbered captions in, one label track per category

Press **Ctrl+Shift+Alt+I** while Audacity is focused. Pick a caption file and it
becomes one label track per category number.

The annotation site hands out files like `Hut R0_sound_nouns.txt`, tab separated,
where a number before the dash says which category the caption belongs to, and an
optional fourth field carries that category's track name:

```text
0.000	30.000	8 – Insect chirping	Insect Chirping
0.200	0.933	1 – Male speech	Male Speech
1.767	4.767	2 – Breathing	Breathing
```

Audacity's own *File > Import > Labels* flattens all of that into a **single**
track, losing the grouping. This builds a track per number instead — named
`1 Male Speech`, `2 Breathing`, `8 Insect Chirping`, in ascending number order
beneath the audio — and strips the `N – ` prefix from each caption, since the
track already carries the number.

Files without that fourth field are the older shape and still import; their
tracks are named `1`, `2`, `8` as before. Two things stop rather than guess: a
category whose rows disagree about the name (`Male speech` and `Male Speech` are
a disagreement, and the file says which lines), and a line carrying a fifth
field. One name used by two categories only warns — `1 Speech` and `4 Speech`
are still distinct tracks.

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
coming back as `1 Male Speech`, `2 Breathing` and so on, or as bare `1`, `2`,
`3` from a file with no names in it.

Overlapping captions inside one category are kept as they are; Audacity label
tracks support them and they occur in real files.

## 3b. Workbook in, label tracks out — one press

Press **Ctrl+Shift+Alt+N** while Audacity is focused. Pick the annotation
`.xlsx`, and the label tracks appear — no caption file to make first, nothing to
download, nothing to drag anywhere.

It reads the sheet, asks Claude for one sound-event noun per caption, builds the
numbered caption file from those nouns, and hands it to the import in
[section 3](#3-numbered-captions-in-one-label-track-per-category). The file is
kept, next to the workbook and named after it — `POW R0.xlsx` sheet `Refinement`
becomes `POW R0 - Refinement Sound Nouns.txt` — so it can be re-imported later
without asking again.

**Claude is asked for the noun and nothing else.** Start, end, the category
number and the track name are copied out of the workbook by the script, so they
cannot come back wrong. This matters: the files made by hand had fields three and
four swapped, which the import refused outright, and that whole class of mistake
is now impossible rather than merely checked for.

The noun-choosing rules are the ones in `Sound_Noun.md` — the same file used when
this is done through Claude on the web, so there is one copy of them and no
chance of the two drifting.

**Everything checkable is checked before the model is called**: one track number
must carry one track name, every row needs a number and a name, times must parse
and not run backwards. A workbook that breaks any of it says so immediately,
naming the rows, rather than after a minute of waiting. Afterwards the finished
file goes through the same parser the import uses before Audacity is touched at
all.

A workbook holding several annotation sheets — `Completion` and `Refinement`
alongside a `Summary` that isn't one — asks which to use; exactly one match binds
silently.

Existing label tracks still get the replace prompt from
[section 3](#3-numbered-captions-in-one-label-track-per-category) — that guard is
the one thing standing between generated captions and real annotation work, so it
stays.

Reckon on half a minute or so: Excel takes a few seconds to open the workbook and
Claude a few more to answer. A small window says which of the two it is waiting
on.

```text
MakeSoundNouns.ps1 -Path "POW R0.xlsx"     # skip the picker
                   -Sheet Refinement       # skip the sheet question
                   -Model claude-sonnet-5  # default is claude-opus-5
                   -NoImport               # write the .txt and stop
                   -Replace                # answer the replace prompt, unattended
                   -DryRun                 # print the prompt, call nothing
```

Requires the Claude Code CLI on PATH and signed in — `npm install -g
@anthropic-ai/claude-code`, then run `claude` once. It bills against whatever
that login already uses; there is no API key to set up here.

## 4. Filling the annotation website — Caption Filler

Double-click **`CaptionFiller.ahk`**. Pick the workbook, and a small always-on-top
guide appears holding one caption at a time.

This one reads the older headers (`Caption`, `Start Time`, `End Time`, …), which
is now the `Legacy` layout — see [Going back to the old
headers](#going-back-to-the-old-headers). Fresh exports carry the `Soup EE`
headers and belong to Caption Filler II/III below.

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
| `Audacity annotation hotkey.ahk` | Binds Ctrl+Shift+Alt+K, I and N inside Audacity (AutoHotkey v2) |
| `ExportAnnotations.ps1` | Live export: reads the running project, prompts, writes .xlsx |
| `ImportSoundNouns.ps1` | Live import: a caption `.txt` in, one label track per category number |
| `MakeSoundNouns.ps1` | Workbook in, label tracks out: asks Claude for the nouns, then runs the import |
| `Sound_Noun.md` | The noun-choosing rules, shared by the hotkey and by Claude on the web |
| `lib\SoundNouns.ps1` | Parses those caption files and groups them by number, with each group's track name |
| `lib\AnnotationWorkbookReader.ps1` | Reads a workbook back in; finds the annotation sheet and its leftmost header block |
| `lib\ClaudeCli.ps1` | Runs one prompt through the Claude Code CLI and hands back the reply |
| `lib\Dialogs.ps1` | The message boxes and the progress window, and the focus fix that lets a hidden process show them |
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
[Ctrl+Shift+Alt+N](#3b-workbook-in-label-tracks-out--one-press) additionally
needs the Claude Code CLI on PATH and signed in; the other two routes do not.

Note: Audacity 4 is also installed here, but it has no scripting pipe — the
hotkey route works with **Audacity 3.7.8** only.

## Changing the hotkeys

Edit the `^+!k::` (export), `^+!i::` (import) or `^+!n::` (make nouns and import)
line in `Audacity annotation hotkey.ahk`.
`^` = Ctrl, `+` = Shift, `!` = Alt.
