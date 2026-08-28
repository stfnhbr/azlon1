# azlon1

Excel VBA macros for tidying caption exports and workbooks.
Import a module into the workbook's VBA project (*VBE → File → Import File…*)
and run the macro of the same name.

## `src/HYX.bas`

Excel VBA macro that tidies a caption export and prefixes it with a
**Caption Number** column.

Run `HYX` on the sheet holding the export. What it does:

1. sets **every worksheet in the workbook** to Consolas 8 pt, top-aligned
   (horizontal alignment reset to General, so text sits left and numbers sit
   right), and unmerges the export sheet;
2. inserts column A, `Caption Number` — unless the sheet already has one, so a
   second run renumbers in place instead of inserting a duplicate column;
3. numbers the caption rows that carry data, and only those: blank rows stay
   blank, and nothing below the last used row is touched;
4. centres column A (Caption Number) and column B (Track Number), and formats
   Start/End time to three decimals;
5. bolds and highlights the header row yellow across the used columns,
   reapplies the AutoFilter across the whole table so the new column gets a
   dropdown too, and freezes the header row.

Only the export sheet is unmerged, and protected sheets are left alone — the
active sheet being protected stops the run with a message.

The raw export's column positions live in the constants at the top of the
module (`RAW_TRACK_COL`, `RAW_TIME_FIRST`, `RAW_TIME_LAST`); adjust those if
the export layout changes.

## `src/CC.bas`

Applies the house formatting to **every worksheet** in the active workbook.

Run `CC` with the workbook open. Per sheet:

1. Consolas 8 pt, top-aligned, across the whole sheet (horizontal alignment
   is reset to General, so text sits left and numbers sit right);
2. where the sheet has column headers — i.e. row 1 is not blank — the header
   row is bolded and highlighted yellow across the used columns, an AutoFilter
   is applied over it, and row 1 is frozen.

Sheets without headers are still formatted, but get no highlight, filter or
freeze. Sheets holding an Excel Table keep the table's own filter rather than
gaining a second, conflicting one. Protected sheets are skipped and listed at
the end of the run. Merged cells are left as they are, since `CC` runs over
whole workbooks and unmerging would wreck any sheet that is laid out rather
than tabular.
