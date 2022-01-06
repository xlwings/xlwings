"""Generates the JSON structure as returned by Excel online"""

import json

import xlwings as xw

book = xw.Book('web.xlsx')

data = {
    'version': 'dev',
    'book': {'name': book.name, 'active_sheet_index': book.sheets.active.index - 1},
    'sheets': []
}

for sheet in book.sheets:
    data['sheets'].append(
        {
            'name': sheet.name,
            'values': [['' if v is None else v for v in row] for row in sheet.used_range.value]
        }
    )

print(json.loads(json.dumps(data, default=lambda d: d.isoformat() + '.000Z')))
