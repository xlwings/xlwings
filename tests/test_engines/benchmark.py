# Make sure to build the Rust library with the release flag
from pathlib import Path

import openpyxl  # noqa: F401
import pandas as pd
import pyxlsb  # noqa: F401
import xlrd  # noqa : F401
from codetiming import Timer

import xlwings as xw

this_dir = Path(__file__).resolve().parent
test_file = this_dir / "single_sheet_many_rows.xlsx"

if test_file.suffix in [".xlsx", ".xlsm"]:
    engine = "openpyxl"
elif test_file.suffix == ".xls":
    # xcalamine doesn't yet support datetime conversion, while xlrd does
    engine = "xlrd"
elif test_file.suffix == ".xlsb":
    # Both, xcalamine and pyxlsb don't yet support datetime conversion
    engine = "pyxlsb"
else:
    engine = None

with Timer(text=f"df1: pandas[{engine}]" + " whole sheet: {:.4f}s"):
    df1 = pd.read_excel(test_file, sheet_name=0, engine=engine)

with Timer(text="df2: xlwings[raw_values] whole sheet: {:.4f}s"):
    with xw.Book(test_file, mode="r") as book:
        data = book.sheets[0].cells.raw_value
    df2 = pd.DataFrame(data=data[1:], columns=data[0])

with Timer(text="df3: xlwings[converter] whole sheet: {:.4f}s"):
    with xw.Book(test_file, mode="r") as book:
        df3 = book.sheets[0].cells.options("df", index=False).value

with Timer(text="df4: xlwings[expand] whole sheet: {:.4f}s"):
    with xw.Book(test_file, mode="r") as book:
        df4 = book.sheets[0]["A1"].expand().options("df", index=False).value

with Timer(text=f"df5: pandas[{engine}]" + "small range: {:.4f}s"):
    df5 = pd.read_excel(test_file, sheet_name=0, nrows=10, engine=engine)

with Timer(text="df6: xlwings[converter] small range: {:.4f}s"):
    with xw.Book(test_file, mode="r") as book:
        df6 = book.sheets[0]["A1:G11"].options("df", index=False).value

print(f"df1 equals df2? {df1.equals(df2)}")
print(f"df1 equals df3? {df1.equals(df3)}")
print(f"df1 equals df4? {df1.equals(df4)}")
print(f"df5 equals df6? {df5.equals(df6)}")
