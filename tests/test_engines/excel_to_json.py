"""Generates the JSON structure as returned by Excel online"""

import json

import xlwings as xw

book = xw.books.active

data = {
    "client": "Microsoft Office Scripts",
    "version": "dev",
    "book": {
        "name": book.name,
        "active_sheet_index": book.sheets.active.index - 1,
        "selection": book.selection.address.replace("$", ""),
    },
    "sheets": [],
}

for sheet in book.sheets:
    last_cell = sheet.used_range.last_cell
    data["sheets"].append(
        {
            "name": sheet.name,
            "values": [
                ["" if v is None else v for v in row]
                for row in sheet[0 : last_cell.row, 0 : last_cell.column].value
            ],
        }
    )

print(json.loads(json.dumps(data, default=lambda d: d.isoformat() + ".000Z")))
