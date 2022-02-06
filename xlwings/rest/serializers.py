def serialize_app(app):
    return {
        "version": str(app.version),
        "visible": app.visible,
        "screen_updating": app.screen_updating,
        "display_alerts": app.display_alerts,
        "calculation": app.calculation,
        "selection": app.selection.get_address(external=True)
        if app.selection
        else None,
        "books": [book.fullname for book in app.books],
        "pid": app.pid,
    }


def serialize_book(book):
    return {
        "app": book.app.pid,
        "name": book.name,
        "sheets": [sheet.name for sheet in book.sheets],
        "fullname": book.fullname,
        "names": [name.name for name in book.names],
        "selection": book.selection.get_address(include_sheetname=True)
        if book.selection
        else None,
    }


def serialize_sheet(sheet):
    return {
        # Omitting index as it is currently 1-based in the Python API
        # but everything in the REST API is 0-based
        "name": sheet.name,
        "names": [name.name for name in sheet.names],
        "charts": [chart.name for chart in sheet.charts],
        "shapes": [shape.name for shape in sheet.shapes],
        "pictures": [picture.name for picture in sheet.pictures],
        "used_range": sheet.used_range.address,
    }


def serialize_chart(chart):
    return {
        "name": chart.name,
        "chart_type": chart.chart_type,
        "height": chart.height,
        "left": chart.left,
        "top": chart.top,
        "width": chart.width,
    }


def serialize_picture(picture):
    return {
        "name": picture.name,
        "height": picture.height,
        "left": picture.left,
        "top": picture.top,
        "width": picture.width,
    }


def serialize_shape(shape):
    return {
        "name": shape.name,
        "type": shape.type,
        "height": shape.height,
        "left": shape.left,
        "top": shape.top,
        "width": shape.width,
    }


def serialize_name(name):
    return {
        "name": name.name,
        "refers_to": name.refers_to,
    }


def serialize_range(rng):
    return {
        "value": rng.value,
        "count": rng.count,
        "row": rng.row,
        "column": rng.column,
        "formula": rng.formula,
        "formula_array": rng.formula_array,
        "column_width": rng.column_width,
        "row_height": rng.row_height,
        "address": rng.address,
        "color": rng.color,
        "current_region": rng.current_region.address,
        "height": rng.height,
        "last_cell": rng.last_cell.address,
        "left": rng.left,
        "name": rng.name.name if rng.name else None,
        "number_format": rng.number_format,
        "shape": rng.shape,
        "size": rng.size,
        "top": rng.top,
        "width": rng.width,
    }
