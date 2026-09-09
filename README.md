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

On **Audacity 3.7.9** this route reads Audacity directly through
`mod-script-pipe`, so there is no Export Labels step — and unlike Audacity's own
`.txt` export, it keeps the name of the label track each annotation came from.

That requires `mod-script-pipe` set to **Enabled** under
*Edit > Preferences > Modules* (already the case here), and a restart of
Audacity after changing that setting.

The same key works in **Audacity 4.0**, which has no pipe to speak to; it reads
the project file instead, and keeps the track names just the same. See
[Audacity 4.0](#6-audacity-40--the-same-keys-a-different-way-in).

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
| 1 | Caption Number | generated — 1..n down the sheet, or the number the labels carry where every one of them has it (see [3b](#3b-workbook-in-label-tracks-out--one-press)) |
| 2 | TrackNumber | Audacity — the label track's position (hotkey route only) |
| 3 | Track | Audacity — e.g. `3 Bird Sounds`, whatever the track is called (hotkey route only). Tracks built by [the import](#3-numbered-captions-in-one-label-track-per-category) are named `<number> <track name>`, so a name put in there comes back out here |
| 4 | TrackDescription | *empty, for you* |
| 5 | StartTime(s) | Audacity |
| 6 | EndTime(s) | Audacity |
| 7 | SourceVisibility | *empty, for you* |
| 8 | SourceDescription | *empty, for you* |
| 9 | Prominence | *empty, for you* |
| 10 | AnnotationText | Audacity — the label text, with the caption number split off into column 1 where there is one |

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

If the project already has label tracks, it asks once what the incoming ones
should do:

| Answer | What happens |
|---|---|
| **Yes** | Replace: the label tracks already there are deleted first |
| **No** | Add alongside: they stay, and the new tracks land beneath them |
| **Cancel** | Nothing at all |

**No** is the default, so Enter cannot delete annotation work by accident, and
Escape still cancels outright. Either import is a single undo step.

Adding alongside is the way to build a project up from more than one caption
file, or to compare two passes over the same audio side by side. The category
numbering restarts per import, so a second import of seven categories gives a
second `1 Male Speech`, `2 Breathing`, … underneath the first — the tracks are
distinct, they simply share names.

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

**Each label carries its caption number**, so a label heard in Audacity can be
found again in the sheet — caption 12 reads `12 Insect chirping`, on a track
still called `8 Insect Chirping`. The number comes straight out of the workbook's
Caption Number column, not from Claude, and `-NoCaptionNumbers` leaves it off.

It also survives the way back: Ctrl+Shift+Alt+K splits the number off the text
again and puts it in column 1, so a workbook that goes out to Audacity and comes
back keeps its own numbering instead of being renumbered down the page. That
happens only when *every* label carries one — a single label typed by hand is
enough to fall back to plain 1..n for the lot, since half a numbering is worse
than none.

**Claude is asked for the noun and nothing else.** Start, end, the category
number and the track name are copied out of the workbook by the script, so they
cannot come back wrong. This matters: the files made by hand had fields three and
four swapped, which the import refused outright, and that whole class of mistake
is now impossible rather than merely checked for.

The noun-choosing rules are the ones in `Sound_Noun.md` — the same file used when
this is done through Claude on the web, so there is one copy of them and no
chance of the two drifting.

**Everything checkable is checked before the model is called**: one track number
must carry one track name, a named track needs a number, times must parse and
not run backwards. A workbook that breaks any of it says so immediately, naming
the rows, rather than after a minute of waiting. Afterwards the finished file
goes through the same parser the import uses before Audacity is touched at all.

**Captions naming no track are kept, not refused.** A blank `Track` cell used to
stop the whole run for want of anything to call the track. They are now gathered
onto one track of their own called `No Track`, numbered after the highest the
sheet uses so it lands at the bottom of the project — work nobody has assigned
yet is still worth hearing in place. Fill the cells in later and the next run
files those captions normally, since a track exported as `9 No Track` reads back
as an ordinary number and name.

**The annotation sheet is found for you.** A workbook holding several — a
`Completion` pass and a `Refinement` pass, alongside a `Summary` that carries no
captions — asks which to use, listing each with the number of captions on it,
since the names alone rarely say which pass is the finished one. Exactly one
match binds silently, and a sheet with the columns but nothing underneath them is
not a match, so a pass that has been prepared and not yet filled in never becomes
a question.

Existing label tracks still get the replace / add-alongside / cancel prompt from
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
                   -NoCaptionNumbers       # labels read "Insect chirping", not "12 Insect chirping"
                   -NoImport               # write the .txt and stop
                   -Replace                # existing label tracks: replace them
                   -KeepExisting           # existing label tracks: add alongside
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

## 6. Audacity 4.0 — the same keys, a different way in

Audacity 4.0 ships **no scripting interface at all**: no `mod-script-pipe`, no
`modules` folder, nothing to send a command to. Everything the three hotkeys did
was built on that pipe, so on 4.0 none of it could work as written.

What 4.0 did keep is the file format. An `.aup4` is a SQLite database with the
same `project` and `autosave` tables an `.aup3` has, holding the same
`ProjectSerializer` binary XML, down to the same `<labeltrack name=…>` with the
same `<label t= t1= title=>` inside it. So the project file is the way in, and
because the format did not change it is one way in for both versions.

The keys are unchanged. Which Audacity you pressed them in decides the route,
and AutoHotkey passes that along so nothing has to be guessed:

| | Audacity 3.7.9 | Audacity 4.0 |
|---|---|---|
| **Ctrl+Shift+Alt+K** export | over the pipe, live | reads the project file |
| **Ctrl+Shift+Alt+I** import | over the pipe, live | writes the project file |
| **Ctrl+Shift+Alt+N** nouns | over the pipe, live | writes the project file |

Two things are worth knowing before you press a key in 4.0.

**Reading is live, not stale.** 4.0 keeps unsaved work in the project's
`autosave` table — and for a project never yet saved, in an `.aup4unsaved` file
under `%LOCALAPPDATA%\Audacity\Audacity4\SessionData`. The export reads whichever
of those is current, so it sees the labels on screen, not the last save.

**Writing needs the project closed.** Labels are written straight into the file,
and Audacity would write its own picture of the project back over them on its
next save. So: save and close the project in 4.0, press the key, and answer Yes
when it offers to open the project again. The scripts check, and refuse with
that explanation rather than writing into a file Audacity is holding.

Every write copies the project to `NAME (before labels 2026-09-09 143811).aup4`
first, and reads the file back off the disk and counts the tracks before calling
it a success.

If both versions are open at once and you run a script by hand rather than by
hotkey, it will not guess — pass `-Version 3` or `-Version 4`.

The originals are untouched and still work: `ExportAnnotations.ps1`,
`ImportSoundNouns.ps1` and `MakeSoundNouns.ps1` are still the 3.7.9 route, and
are what the new scripts call for it.

## Files

| File | Role |
|---|---|
| `Audacity annotation hotkey.ahk` | Binds Ctrl+Shift+Alt+K, I and N inside either Audacity, and Ctrl+Alt+/ anywhere (AutoHotkey v2) |
| `ExportAnnotations2.ps1` | Export for both versions: pipe on 3.7.9, project file on 4.0 |
| `MakeSoundNouns2.ps1` | Workbook or caption file in, label tracks out, on both versions |
| `lib\AudacityProjectFile.ps1` | Reads and writes `.aup3`/`.aup4` directly: SQLite, and the binary XML inside it |
| `lib\AudacityApps.ps1` | Which Audacity is running and in front, and which project file that means |
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

Excel (uses COM to write real `.xlsx`, not CSV), AutoHotkey v2, and Audacity —
**3.7.9** with `mod-script-pipe`, or **4.0**, or both. All present on this
machine. [Ctrl+Shift+Alt+N](#3b-workbook-in-label-tracks-out--one-press)
additionally needs the Claude Code CLI on PATH and signed in; the other two
routes do not.

The 4.0 route needs nothing extra at all: it reads and writes the project file
through `winsqlite3.dll`, which ships with Windows.

## Changing the hotkeys

Edit the `^+!k::` (export), `^+!i::` (import), `^+!n::` (make nouns and import)
or `^!/::` (open Caption Filler III) line in `Audacity annotation hotkey.ahk`.
`^` = Ctrl, `+` = Shift, `!` = Alt.

The first three fire only while Audacity is focused — either version, and which
one you pressed them in is passed to the script as `-Version`. **Ctrl+Alt+/** fires
anywhere, because the guide it opens belongs beside the browser — pressing it
again while the guide is up raises that window rather than starting a second
copy, so the sheet stays bound and the `✓` marks stay put. It is the `/` key of
the US layout; on the Danish, Norwegian and Swedish layouts also installed here
`/` is Shift+7, so write `^!SC035::` instead to keep the same physical key.
