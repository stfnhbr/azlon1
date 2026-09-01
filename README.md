# azlon1

Excel macro that renumbers the tracks in a captioning workbook so track numbers
run 1..N in the chronological order of each track's first caption.

## The problem

On the `Refinement` sheet, `TrackNumber` values are carried over from the
earlier annotation pass, so they no longer follow the order in which the tracks
are actually first heard. In `IGN_R0.xlsx` the column reads
1, 2, 9, 8, 3, 4, 6, 6, 4, 6, 9, 7, 6, 5, 7, 6, 10 — track 9 (Bird calls) first
occurs at 0.5s but is numbered after track 8, and so on.

## What the macro does

For every track it finds the earliest `StartTime(s)` across that track's caption
rows, sorts the tracks by that time, and writes back 1..N in that order. On the
sample workbook:

| old | new | first caption | track |
|----:|----:|--------------:|-------|
|  1 |  1 |  0.000s | Background music |
|  2 |  2 |  0.000s | Jogger's footsteps |
|  9 |  3 |  0.500s | Bird calls |
|  8 |  4 |  1.233s | Ambient atmospheric noise |
|  3 |  5 |  1.467s | Background chatter |
|  4 |  6 |  2.300s | Off-screen footsteps |
|  6 |  7 |  2.667s | Unidentified background noises |
|  7 |  8 | 10.467s | Engine rumbling |
|  5 |  9 | 10.733s | Male voice |
| 10 | 10 | 13.467s | Female voice |

Only column B changes; caption rows keep their existing order and every other
column is left alone.

## Installing it

1. Open the workbook and press <kbd>Alt</kbd>+<kbd>F11</kbd> (Mac:
   <kbd>Fn</kbd>+<kbd>Option</kbd>+<kbd>F11</kbd>) to open the VBA editor.
2. **File → Import File…** and pick `vba/modTrackNumbering.bas`.
3. Back in Excel, **Developer → Macros**, choose
   `RenumberTracksByFirstCaption`, and click **Run**. (Or press
   <kbd>Alt</kbd>+<kbd>F8</kbd> if the Developer tab is hidden.)
4. Save as `.xlsm` if you want the macro to stay with the workbook — a plain
   `.xlsx` cannot store macros.

## Macros

| Macro | What it does |
|---|---|
| `RenumberTracksByFirstCaption` | Renumbers the sheet named by `SHEET_NAME` (default `Refinement`). |
| `RenumberTracksOnActiveSheet` | Renumbers whichever sheet is currently active — handy for `Completion` or a differently-named sheet. |
| `UndoLastTrackRenumber` | Restores the numbers from the last run. |

Running any macro clears Excel's undo stack, so <kbd>Ctrl</kbd>+<kbd>Z</kbd>
will **not** undo the renumbering. The macro saves the original column to a
very-hidden `_TrackNumberBackup` sheet first; `UndoLastTrackRenumber` reads it
back.

## Settings

Constants at the top of the module:

| Constant | Default | Purpose |
|---|---|---|
| `SHEET_NAME` | `"Refinement"` | Sheet that `RenumberTracksByFirstCaption` works on. |
| `HEADER_ROW` | `1` | Row holding the headers; data starts on the next row. |
| `GROUP_BY` | `"TrackNumber"` | How rows are grouped into tracks. `"TrackNumber"` groups by the existing number and falls back to the track name where the number is blank; `"Track"` groups by name only — use it when the existing numbers are unreliable. |
| `RESEQUENCE_CAPTION_NUMBERS` | `False` | Set `True` to also renumber the `Caption Number` column 1..N in row order. |
| `SHOW_SUMMARY` | `True` | Show the old → new mapping when the macro finishes. |

## Behaviour worth knowing

- **Columns are found by header text, not by letter.** Inserting or reordering
  columns is safe. Matching ignores case, spaces, underscores and punctuation,
  so `StartTime(s)`, `Start Time (s)` and `start_time` all match.
- **Row order does not matter.** The macro reads each track's earliest start
  time rather than assuming the sheet is already sorted, and it never moves
  rows.
- **Ties break by sheet position.** Background music and Jogger's footsteps both
  start at 0.000s, so the one appearing higher on the sheet takes the lower
  number.
- **Tracks with no usable start time sort last**, in the order they first
  appear.
- **Rows with neither a track number nor a track name are left untouched** and
  counted as skipped in the summary.
- **Running it twice is a no-op** — the second run reports 0 changes.
- No external references are needed (no `Scripting.Dictionary`), so it runs on
  Excel for Windows and Excel for Mac.
