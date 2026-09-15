"""
Open workbooks in ONE Excel instead of starting a new Excel every time.

Why this exists
---------------
Two habits cause the "PERSONAL.XLSB is locked for editing by <you>" prompt:

1. A script creates Excel through COM (win32com or xlwings), shows a file to
   you, and never quits.  That Excel keeps running in the background and holds
   PERSONAL.XLSB open.
2. A second script then "opens" a file with os.startfile(path) or
   subprocess.  Windows starts a brand-new EXCEL.EXE for that, which tries to
   open PERSONAL.XLSB again -> locked.

xlwings fixes both.  xw.Book(path) looks for a running Excel first and opens
the file inside it.  Only when no Excel is running does it start one.

Install once:  pip install xlwings

Usage
-----
    from excel_open import show, work_quietly

    # Case A: open a file so the user can look at it (keeps Excel open)
    show(r"C:\Users\me\OneDrive\Desktop\file.xlsx")

    # Case B: read or change a file in the background, then close Excel
    with work_quietly(r"C:\Users\me\file.xlsx") as wb:
        wb.sheets[0]["A1"].value = "hello"
        wb.save()
"""
from contextlib import contextmanager
from pathlib import Path

import xlwings as xw


def show(path):
    """Open a workbook for the user in the Excel that is already running.

    Never starts a second Excel when one is open.  Returns the xlwings Book.
    Do NOT call app.quit() afterwards, the user is still using it.
    """
    path = str(Path(path).resolve())
    book = xw.Book(path)          # reuses the running Excel, or starts one if none
    book.app.visible = True
    book.activate()
    return book


@contextmanager
def work_quietly(path, save=False):
    """Open a workbook in a hidden, private Excel and ALWAYS close it afterwards.

    Use this for scripts that only read or edit data.  The `finally` block
    guarantees Excel quits even when your code raises an error, so no
    hidden Excel is ever left behind.
    """
    path = str(Path(path).resolve())
    app = xw.App(visible=False, add_book=False)
    try:
        book = app.books.open(path)
        yield book
        if save:
            book.save()
        book.close()
    finally:
        app.quit()


if __name__ == "__main__":
    import sys
    if len(sys.argv) != 2:
        print("Usage: python excel_open.py <path-to-workbook>")
        sys.exit(1)
    show(sys.argv[1])
