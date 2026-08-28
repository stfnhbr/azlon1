# azlon1

## `src/HYX.bas`

Excel VBA macro that tidies a caption export and prefixes it with a
**Caption Number** column.

Import `src/HYX.bas` into the workbook's VBA project
(*VBE → File → Import File…*), then run `HYX` on the sheet holding the
export.

What it does:

1. normalises sheet-wide formatting (Consolas 8, top-aligned, no merged cells);
2. inserts column A, `Caption Number` — unless the sheet already has one, so a
   second run renumbers in place instead of inserting a duplicate column;
3. numbers the caption rows that carry data, and only those: blank rows stay
   blank, and nothing below the last used row is touched;
4. centres Caption Number and Track Number, and formats Start/End time to
   three decimals;
5. reapplies the AutoFilter across the whole table so the new column gets a
   dropdown too.

The raw export's column positions live in the constants at the top of the
module (`RAW_TRACK_COL`, `RAW_TIME_FIRST`, `RAW_TIME_LAST`); adjust those if
the export layout changes.
