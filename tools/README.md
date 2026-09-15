# tools

## Fix-PersonalXlsb.ps1

One-shot fix for the Excel prompt
"PERSONAL.XLSB is locked for editing by <your name>".

1. Save any open Excel work. The script closes Excel without asking.
2. Open PowerShell (Start menu, type `powershell`, press Enter).
3. Run:

```powershell
powershell -ExecutionPolicy Bypass -File "$HOME\Downloads\Fix-PersonalXlsb.ps1"
```

Adjust the path to wherever you saved the script.

What it does, in order:

| Step | Action |
|------|--------|
| 1 | Ends every EXCEL.EXE process, including hidden ones left behind by scripts |
| 2 | Deletes stale `~$PERSONAL.XLSB` lock files in `%APPDATA%\Microsoft\Excel\XLSTART` |
| 3 | Turns off "Ignore other applications that use DDE" in Excel options |
| 4 | Reports duplicate `PERSONAL.XLSB` copies in the Office program startup folders |
| 5 | Warns if the startup folder is inside OneDrive |
| 6 | Optionally sets the read-only attribute (`-MakeReadOnly`) so Excel never asks again |

Switches:

- `-MakeReadOnly` marks `PERSONAL.XLSB` read-only. The prompt disappears, but you must run
  `-ClearReadOnly` before saving new macros into it.
- `-ClearReadOnly` undoes that.
- `-SkipDde` skips step 3 if you prefer to change the option by hand.

## If you automate Excel from Python

A leftover Excel process is the most common cause. Always quit Excel in a `finally` block:

```python
import xlwings as xw

app = xw.App(visible=False)
try:
    wb = app.books.open(r"C:\path\to\file.xlsx")
    # your work here
    wb.save()
finally:
    app.quit()
```

## Find-PersonalMacros.ps1

Use this when the personal macros seem to have disappeared after a crash or a forced
close of Excel. Close Excel first, then run:

```powershell
powershell -ExecutionPolicy Bypass -File "$HOME\Downloads\Find-PersonalMacros.ps1"
```

It reports where `PERSONAL.XLSB` is, lists every copy on your profile, shows Excel's
"Disabled Items" list, and checks whether the file still contains VBA code. Nothing is
changed unless you add `-ReEnable`, which clears the Disabled Items list so Excel loads
the file again at startup.
