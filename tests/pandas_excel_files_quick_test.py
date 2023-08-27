"""
This is just a quick test to see if all Excel files of the pandas test suite can be
opened and all sheets can be read properly.

=> Requires the pandas repo checked out next to the xlwings repo
=> "passes" if no Exception is raised
"""

from pathlib import Path

import xlwings as xw

path = (
    Path(__file__).resolve().parent.parent.parent
    / "pandas"
    / "pandas"
    / "tests"
    / "io"
    / "data"
    / "excel"
)

ix = 0
for ix, f in enumerate(path.glob("[!~$]*.xls*")):
    file_name = f.resolve()
    try:
        with xw.Book(f.resolve(), mode="r") as book:
            for sheet in book.sheets:
                sheet_name = sheet.name
                data = sheet.cells.value
        print(".", end="", flush=True)
    except:  # noqa: E722
        print(f"FAILED: {f.resolve()} [repr(e)]")

print()
print(f"Tested {ix + 1} files.")
