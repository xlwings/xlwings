import logging
import xlwings as xw

from flask import Flask, jsonify


app = Flask(__name__)
logger = logging.getLogger()


def serialize_worksheet(sheet):
    return {
        'id': None,
        'position': sheet.index,
        'name': sheet.name,
        'visibility': 'Visible' if sheet.api.Visible == -1 else None
    }

def get_sheet(workbook, name_or_id):
    if name_or_id.isdigit():
        name_or_id = int(name_or_id)
    return workbook.sheets[name_or_id]


# GET https://graph.microsoft.com/v1.0/me/drive/items/{id}/workbook/worksheets
@app.route('/<path:path>/workbook/worksheets', methods=['GET'])
def worksheets(path):
    # TODO: POST (create)
    workbook = xw.Book(path)
    return jsonify(value=[serialize_worksheet(worksheet) 
        for worksheet in workbook.sheets])


# https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/worksheet_get
# GET https://graph.microsoft.com/v1.0/me/drive/items/{id}/workbook/worksheets/{id|name}
@app.route('/<path:path>/workbook/worksheets/<string:name_or_id>', methods=['GET'])
def worksheet(path, name_or_id):
    # TODO: DELETE (delete)
    workbook = xw.Book(path)
    if name_or_id.isdigit():
        name_or_id = int(name_or_id) - 1 
    worksheet = workbook.sheets[name_or_id]
    return jsonify(serialize_worksheet(worksheet))


# https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/workbook_list_names
# GET https://graph.microsoft.com/v1.0/me/drive/items/{id}/workbook/names
@app.route('/<path:path>/workbook/names', methods=['GET'])
def names(path):
    workbook = xw.Book(path)
    return jsonify(value=[
        {
            'name': name.name,
            'type': "Range",
            'value': name.refers_to[1:],
            'visible': True
            
        } 
        for name in workbook.names
    ])


# https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/worksheet_range
# GET https://graph.microsoft.com/v1.0/me/drive/items/{id}/workbook/worksheets/{id|name}/range(address='A1:B2')
@app.route('/<path:path>/workbook/worksheets/<string:name_or_id>/range(address=<string:address>)', methods=['GET'])
@app.route('/<path:path>/workbook/worksheets/<string:name_or_id>/range', defaults={'address': None}, methods=['GET'])
def rng(path, name_or_id, address=None):
    workbook = xw.Book(path)
    sheet = get_sheet(workbook, name_or_id)
    if not address:
        address = sheet.api.UsedRange.Address
    rng = sheet[address]
    rows_count = rng.rows.count
    columns_count = rng.columns.count
    return jsonify({
        "address": sheet.name + '!' + rng.address,
        "addressLocal": sheet.name + '!' + rng.address,
        "cellCount": rng.count,
        "columnCount": columns_count,
        "columnHidden": False,
        "columnIndex": rng.columns[0].column,
        "formulas": rng.formula,
        "formulasLocal": rng.formula,  # TODO!
        "formulasR1C1": rng.formula,  # TODO!
        "hidden": False,
        "numberFormat": [[None for i in range(columns_count)] for i in range(rows_count)], #[[cell.number_format for cell in row] for row in rng.rows],  # TODO: is this fast enough for big ranges?
        "rowCount": rows_count,
        "rowHidden": False,
        "rowIndex": rng.rows[0].row,
        "text": rng.value,  # TODO!
        "values": rng.value,
        "valueTypes": [[None for i in range(columns_count)] for i in range(rows_count)] # TODO: just prints None at the moment + is this fast enough for big ranges?
    })


# https://developer.microsoft.com/en-us/graph/docs/api-reference/v1.0/api/worksheet_cell
# GET /workbook/worksheets/{id|name}/cell(row={row},column={column})
@app.route('/<path:path>/workbook/worksheets/<string:name_or_id>/cell(row=<int:row>,column=<int:column>)', methods=['GET'])
def cell(path, name_or_id, row, column):
    workbook = xw.Book(path)
    sheet = get_sheet(workbook, name_or_id)
    rng = sheet.cells[row, column]
    return jsonify({
        "address": rng.address,
        "addressLocal": rng.address,
        "cellCount": rng.count,
        "columnCount": rng.columns.count,
        "columnIndex": rng.columns[0].column,
        "valueTypes": None
    })


if __name__ == '__main__':
    app.run(debug=True)
