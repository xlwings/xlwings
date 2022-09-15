# Make sure to build the Rust library with the release flag
import time
from pathlib import Path

import openpyxl
import xlrd
import pyxlsb
import pandas as pd
import xlwings as xw

this_dir = Path(__file__).resolve().parent
test_file = this_dir / "single_sheet_many_rows.xlsx"

if test_file.suffix in [".xlsx", ".xlsm"]:
    engine = "openpyxl"
elif test_file.suffix == ".xls":
    engine = "xlrd"
elif test_file.suffix == ".xlsb":
    engine = "pyxlsb"
else:
    engine = None

# pandas with OpenPyXL (whole sheet)
start_time = time.time()
df1 = pd.read_excel(test_file, sheet_name=0, engine=engine)
print(f"df1: pandas ({engine}) - whole sheet: {time.time() - start_time}s")

# xlwings with raw_values (whole sheet)
start_time = time.time()
with xw.Book(test_file, mode="r") as book:
    data = book.sheets[0].cells.raw_value
df2 = pd.DataFrame(data=data[1:], columns=data[0])
print(f"df2: xlwings (raw_values) - whole sheet: {time.time() - start_time}s")

# xlwings with converter (whole sheet)
start_time = time.time()
with xw.Book(test_file, mode="r") as book:
    df3 = book.sheets[0].cells.options("df", index=False).value
print(f"df3: xlwings (converter) - whole sheet: {time.time() - start_time}s")

# xlwings with converter (expand)
start_time = time.time()
with xw.Book(test_file, mode="r") as book:
    df4 = book.sheets[0]["A1"].expand().options("df", index=False).value
print(f"df4: xlwings (expand) - whole sheet: {time.time() - start_time}s")

# pandas with OpenPyxl (specific range)
start_time = time.time()
df5 = pd.read_excel(test_file, sheet_name=0, nrows=10, engine=engine)
print(f"df5: pandas ({engine}) - small range: {time.time() - start_time}s")

# xlwings with converter (specific range)
start_time = time.time()
with xw.Book(test_file, mode="r") as book:
    df6 = book.sheets[0]["A1:G11"].options("df", index=False).value
print(f"df6: xlwings (converter) - small range: {time.time() - start_time}s")

print(f"df1 equals df2? {df1.equals(df2)}")
print(f"df1 equals df3? {df1.equals(df3)}")
print(f"df1 equals df4? {df1.equals(df4)}")
print(f"df5 equals df6? {df5.equals(df6)}")
