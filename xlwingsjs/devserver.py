import datetime as dt
from pathlib import Path

import custom_functions
import jinja2
import markupsafe
from dateutil import tz
from fastapi import Body, FastAPI, Request, status
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import HTMLResponse, PlainTextResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

import xlwings as xw

# from tests import udf_tests_officejs as custom_functions

app = FastAPI()

this_dir = Path(__file__).resolve().parent


@app.post("/hello")
def hello(request: Request, data: dict = Body):
    print(request.headers)
    book = xw.Book(json=data)
    sheet = book.sheets[0]
    cell = sheet["A1"]
    if cell.value == "Hello xlwings!":
        cell.value = "Bye xlwings!"
    else:
        cell.value = "Hello xlwings!"
    return book.json()


@app.post("/show-alert")
def show_alert(data: dict = Body):
    book = xw.Book(json=data)
    book.app.alert(
        prompt="Some text",
        title="Some Title",
        buttons="ok_cancel",
        mode="info",
        callback="myCallback",
    )
    return book.json()


@app.post("/integration-test-read")
def integration_test_read(data: dict = Body):
    book = xw.Book(json=data)
    assert book.name == "engines.xlsm", "engines.xlsm must be the active file"
    if data["client"] == "Office.js":
        expected_data = expected_body["Office.js"]
    elif data["client"] == "VBA":
        expected_data = expected_body["VBA"]
    elif data["client"] == "Google Apps Script":
        expected_data = expected_body["Google Apps Script"]
    elif data["client"] == "Microsoft Office Scripts":
        expected_data = expected_body["Office Scripts"]
    assert data == expected_data, "Body differs (Make sure to select cell 'Sheet 1'!A1)"
    book.app.alert("OK", title="Integration Test Read")
    return book.json()


@app.post("/integration-test-write")
def integration_test_write(data: dict = Body):
    book = xw.Book(json=data)
    assert (
        book.name == "integration_write.xlsm"
    ), "integration_write.xlsm must be the active file"
    sheet1 = book.sheets["Sheet 1"]

    # Values
    sheet1["E3"].value = [
        [None, "string"],
        [-1, 1],
        [-1.1, 1.1],
        [True, False],
        [
            dt.date(2021, 7, 1),
            dt.datetime(2021, 12, 31, 23, 35, 12, tzinfo=tz.gettz("Europe/Paris")),
        ],
    ]

    # Add sheets and write to them
    # NOTE: in Excel Online, adding/renaming sheets makes an alert impossible to quit
    # via provided buttons ("osfControl for the given ID doesn't exist.")
    sheet2 = book.sheets.add("New Named Sheet")
    sheet2["A1"].value = "Named Sheet"
    sheet3 = book.sheets.add()
    sheet3["A1"].value = "Unnamed Sheet"

    # Tables
    if data["client"] != "Google Apps Script":
        sheet_tables = book.sheets["Tables"]

        sheet_tables["A1"].value = [["one", "two"], [1, 2], [3, 4]]
        sheet_tables.tables.add(sheet3["A1:B3"])

        sheet_tables["A5"].value = [[1, 2], [3, 4]]
        sheet_tables.tables.add(sheet_tables["A5:B6"], has_headers=False)

        sheet_tables["A9"].value = [["one", "two"], [1, 2], [3, 4]]
        mytable1 = sheet_tables.tables.add(sheet_tables["A9:B11"], name="MyTable1")
        mytable1.show_autofilter = False

        sheet_tables["A13"].value = [[1, 2], [3, 4]]
        mytable2 = sheet_tables.tables.add(
            sheet_tables["A13:B14"], name="MyTable2", has_headers=False
        )
        mytable2.show_headers = False
        mytable2.show_totals = True
        mytable2.show_filters = False
        mytable2.resize(sheet_tables["A14:C17"])

    # Set sheet name
    book.sheets["Sheet2"].name = "Changed"
    book.sheets["Changed"]["A1"].value = "Changed"

    # Autofit
    autofit_sheet = book.sheets["Autofit"]
    autofit_sheet["A1"].value = [[1, 2], [3, 4]]
    autofit_sheet["A1:B2"].autofit()
    autofit_sheet["D4:E5"].rows.autofit()
    autofit_sheet["G7:H82"].columns.autofit()

    # Range color
    sheet1["E12:F12"].color = "#3DBAC1"

    # Add Hyperlink
    sheet1["E14"].add_hyperlink("https://www.xlwings.org", "xlwings", "xw homepage")

    # Number format
    sheet1["E9:F10"].value = [[1, 2], [3, 4]]
    sheet1["E9:F10"].number_format = "0%"

    # Clear contents
    sheet1["E16:F17"].clear_contents()

    # Activate sheet
    book.sheets[1].activate()

    # Pictures
    sheet1.pictures.add(
        this_dir.parent / "tests" / "sample_picture.png",
        name="MyPic",
        anchor=sheet1["C28"],
    )
    sheet1.pictures.add(this_dir / "icons" / "icon-80.png", name="MyPic", update=True)
    book.sheets[1].pictures.add(this_dir.parent / "tests" / "sample_picture.png")

    # Add named ranges
    book.names.add("test1", "='Sheet 1'!$A$1:$B$3")
    book.names.add("test2", "=Changed!$A$1")
    sheet1["A1"].name = "test3"
    if data["client"] != "Google Apps Script":
        sheet1.names.add("test4", "='Sheet 1'!$A$1:$B$3")

    # Delete named ranges
    book.names["DeleteMe"].delete()
    book.names["Sheet4!DeleteMe"].delete()
    book.names["'Sheet 3'!DeleteMe"].delete()

    return book.json()


@app.get("/xlwings/alert", response_class=HTMLResponse)
async def alert(
    request: Request, prompt: str, title: str, buttons: str, mode: str, callback: str
):
    """This endpoint is required by myapp.alert()"""
    return templates.TemplateResponse(
        "xlwings-alert.html",
        {
            "request": request,
            "prompt": markupsafe.escape(prompt).replace(
                "\n", markupsafe.Markup("<br>")
            ),
            "title": title,
            "buttons": buttons,
            "mode": mode,
            "callback": callback,
        },
    )


@app.get("/xlwings/custom-functions-meta")
async def custom_functions_meta():
    return xw.server.custom_functions_meta(custom_functions)


@app.get("/xlwings/custom-functions-code")
async def custom_functions_code():
    return PlainTextResponse(xw.server.custom_functions_code(custom_functions))


@app.post("/xlwings/custom-functions-call")
async def custom_functions_call(request: Request, data: dict = Body):
    print(request.headers["Authorization"])
    rv = await xw.server.custom_functions_call(data, custom_functions)
    return {"result": rv}


app.mount("/icons", StaticFiles(directory=this_dir / "icons"), name="icons")
app.mount("/", StaticFiles(directory=this_dir / "build"), name="home")
StaticFiles.is_not_modified = lambda *args, **kwargs: False  # Never cache static files

# Add the xlwings alert template as source
loader = jinja2.ChoiceLoader(
    [
        jinja2.FileSystemLoader(this_dir / "build"),
        jinja2.PackageLoader("xlwings", "html"),
    ]
)
templates = Jinja2Templates(directory=this_dir / "build", loader=loader)


# Office Scripts requires CORS and the following would be enough:
# allow_origin_regex=r"https://.*.officescripts.microsoftusercontent.com"
# but Office.js from Excel on the web also requires CORS and will have other origins
app.add_middleware(
    CORSMiddleware,
    allow_origins="*",
    allow_methods=["POST"],
    allow_headers=["*"],
)


@app.exception_handler(Exception)
async def exception_handler(request, exception):
    # Handling all Exceptions is OK since it's only a dev server, but you probably
    # don't want to show the details of every Exception to the user in production
    return PlainTextResponse(
        repr(exception), status_code=status.HTTP_500_INTERNAL_SERVER_ERROR
    )


expected_body = {}
expected_body["Office.js"] = {
    "client": "Office.js",
    "version": "dev",
    "book": {"name": "engines.xlsm", "active_sheet_index": 0, "selection": "A1"},
    "names": [
        {
            "name": "one",
            "sheet_index": 0,
            "address": "A1",
            "scope_sheet_name": None,
            "scope_sheet_index": None,
            "book_scope": True,
        },
        {
            "name": "two",
            "sheet_index": 1,
            "address": "A1:A2",
            "scope_sheet_name": None,
            "scope_sheet_index": None,
            "book_scope": True,
        },
        {
            "name": "two",
            "sheet_index": 0,
            "address": "C7:D8",
            "scope_sheet_name": "Sheet 1",
            "scope_sheet_index": 0,
            "book_scope": False,
        },
        {
            "name": "two",
            "sheet_index": 2,
            "address": "B3",
            "scope_sheet_name": "Sheet2",
            "scope_sheet_index": 1,
            "book_scope": False,
        },
    ],
    "sheets": [
        {
            "name": "Sheet 1",
            "values": [
                ["a", "b", "c", ""],
                [1.1, 2.2, 3.3, "2021-01-01T00:00:00.000Z"],
                [4.4, 5.5, 6.6, ""],
                ["", "", "", ""],
                ["", "", "", ""],
                ["", "", "", ""],
                ["", "", "", ""],
                ["", "", "", ""],
                ["", "", "", ""],
                ["Column1", "Column2", "", ""],
                [1.1, 2.2, "", ""],
                ["", "", "", ""],
                ["", "", "", ""],
                ["", "", "", ""],
                [1.1, 2.2, 3.3, ""],
                [4.4, 5.5, 6.6, ""],
                ["Total", "", 9.9, ""],
            ],
            # Width/height differ for Desktop and Web
            "pictures": [
                {"name": "mypic1", "height": 10, "width": 20},
                {"name": "mypic2", "height": 30, "width": 40},
            ],
            "tables": [
                {
                    "name": "Table1",
                    "range_address": "A10:B11",
                    "header_row_range_address": "A10:B10",
                    "data_body_range_address": "A11:B11",
                    "total_row_range_address": None,
                    "show_headers": True,
                    "show_totals": False,
                    "table_style": "TableStyleMedium2",
                    "show_autofilter": True,
                },
                {
                    "name": "Table2",
                    "range_address": "A15:C17",
                    "header_row_range_address": None,
                    "data_body_range_address": "A15:C16",
                    "total_row_range_address": "A17:C17",
                    "show_headers": False,
                    "show_totals": True,
                    "table_style": "TableStyleLight1",
                    "show_autofilter": False,
                },
            ],
        },
        {
            "name": "Sheet2",
            "values": [["aa", "bb"], [11.1, 22.2]],
            "pictures": [],
            "tables": [],
        },
        {
            "name": "Sheet3",
            "values": [
                ["", "string"],
                [-1.1, 1.1],
                [True, False],
                ["2021-10-01T00:00:00.000Z", 44561.9826388889],
            ],
            "pictures": [],
            "tables": [],
        },
    ],
}

expected_body["VBA"] = {
    "client": "VBA",
    "version": "dev",
    "book": {"name": "engines.xlsm", "active_sheet_index": 0, "selection": "A1"},
    "names": [
        {
            "name": "one",
            "sheet_index": 0,
            "address": "A1",
            "book_scope": True,
            "scope_sheet_name": None,
            "scope_sheet_index": None,
        },
        {
            "name": "'Sheet 1'!two",
            "sheet_index": 0,
            "address": "C7:D8",
            "book_scope": False,
            "scope_sheet_name": "Sheet 1",
            "scope_sheet_index": 0,
        },
        {
            "name": "Sheet2!two",
            "sheet_index": 2,
            "address": "B3",
            "book_scope": False,
            "scope_sheet_name": "Sheet2",
            "scope_sheet_index": 1,
        },
        {
            "name": "two",
            "sheet_index": 1,
            "address": "A1:A2",
            "book_scope": True,
            "scope_sheet_name": None,
            "scope_sheet_index": None,
        },
    ],
    "sheets": [
        {
            "name": "Sheet 1",
            "pictures": [
                {"name": "mypic1", "height": 10, "width": 20},
                {"name": "mypic2", "height": 30, "width": 40},
            ],
            "tables": [
                {
                    "name": "Table1",
                    "range_address": "$A$10:$B$11",
                    "header_row_range_address": "$A$10:$B$10",
                    "data_body_range_address": "$A$11:$B$11",
                    "total_row_range_address": None,
                    "show_headers": True,
                    "show_totals": False,
                    "table_style": "TableStyleMedium2",
                    "show_autofilter": True,
                },
                {
                    "name": "Table2",
                    "range_address": "$A$15:$C$17",
                    "header_row_range_address": None,
                    "data_body_range_address": "$A$15:$C$16",
                    "total_row_range_address": "$A$17:$C$17",
                    "show_headers": False,
                    "show_totals": True,
                    "table_style": "TableStyleLight1",
                    "show_autofilter": False,
                },
            ],
            "values": [
                ["a", "b", "c", None],
                [1.1, 2.2, 3.3, "2021-01-01T00:00:00.000Z"],
                [4.4, 5.5, 6.6, None],
                [None, None, None, None],
                [None, None, None, None],
                [None, None, None, None],
                [None, None, None, None],
                [None, None, None, None],
                [None, None, None, None],
                ["Column1", "Column2", None, None],
                [1.1, 2.2, None, None],
                [None, None, None, None],
                [None, None, None, None],
                [None, None, None, None],
                [1.1, 2.2, 3.3, None],
                [4.4, 5.5, 6.6, None],
                ["Total", None, 9.9, None],
            ],
        },
        {
            "name": "Sheet2",
            "pictures": [],
            "tables": [],
            "values": [["aa", "bb"], [11.1, 22.2]],
        },
        {
            "name": "Sheet3",
            "pictures": [],
            "tables": [],
            "values": [
                [None, "string"],
                [-1.1, 1.1],
                [True, False],
                ["2021-10-01T00:00:00.000Z", "2021-12-31T23:35:00.000Z"],
            ],
        },
    ],
}
expected_body["Office Scripts"] = {
    "client": "Microsoft Office Scripts",
    "version": "dev",
    "book": {"name": "engines.xlsm", "active_sheet_index": 0, "selection": "A1"},
    "names": [
        {
            "name": "one",
            "sheet_index": 0,
            "address": "A1",
            "scope_sheet_name": None,
            "scope_sheet_index": None,
            "book_scope": True,
        },
        {
            "name": "two",
            "sheet_index": 1,
            "address": "A1:A2",
            "scope_sheet_name": None,
            "scope_sheet_index": None,
            "book_scope": True,
        },
        {
            "name": "two",
            "sheet_index": 0,
            "address": "C7:D8",
            "scope_sheet_name": "Sheet 1",
            "scope_sheet_index": 0,
            "book_scope": False,
        },
        {
            "name": "two",
            "sheet_index": 2,
            "address": "B3",
            "scope_sheet_name": "Sheet2",
            "scope_sheet_index": 1,
            "book_scope": False,
        },
    ],
    "sheets": [
        {
            "name": "Sheet 1",
            "values": [
                ["a", "b", "c", ""],
                [1.1, 2.2, 3.3, "2021-01-01T00:00:00.000Z"],
                [4.4, 5.5, 6.6, ""],
                ["", "", "", ""],
                ["", "", "", ""],
                ["", "", "", ""],
                ["", "", "", ""],
                ["", "", "", ""],
                ["", "", "", ""],
                ["Column1", "Column2", "", ""],
                [1.1, 2.2, "", ""],
                ["", "", "", ""],
                ["", "", "", ""],
                ["", "", "", ""],
                [1.1, 2.2, 3.3, ""],
                [4.4, 5.5, 6.6, ""],
                ["Total", "", 9.9, ""],
            ],
            # Width/height differ for Desktop and Web
            "pictures": [
                {"name": "mypic1", "width": 20, "height": 10},
                {"name": "mypic2", "width": 40, "height": 30},
            ],
            "tables": [
                {
                    "name": "Table1",
                    "range_address": "A10:B11",
                    "header_row_range_address": "A10:B10",
                    "data_body_range_address": "A11:B11",
                    "total_row_range_address": None,
                    "show_headers": True,
                    "show_totals": False,
                    "table_style": "TableStyleMedium2",
                    "show_autofilter": True,
                },
                {
                    "name": "Table2",
                    "range_address": "A15:C17",
                    "header_row_range_address": None,
                    "data_body_range_address": "A15:C16",
                    "total_row_range_address": "A17:C17",
                    "show_headers": False,
                    "show_totals": True,
                    "table_style": "TableStyleLight1",
                    "show_autofilter": False,
                },
            ],
        },
        {
            "name": "Sheet2",
            "values": [["aa", "bb"], [11.1, 22.2]],
            "pictures": [],
            "tables": [],
        },
        {
            "name": "Sheet3",
            "values": [
                ["", "string"],
                [-1.1, 1.1],
                [True, False],
                ["2021-10-01T00:00:00.000Z", 44561.9826388889],
            ],
            "pictures": [],
            "tables": [],
        },
    ],
}
expected_body["Google Apps Script"] = {
    "client": "Google Apps Script",
    "version": "dev",
    "book": {"name": "engines.xlsm", "active_sheet_index": 0, "selection": "A1"},
    "names": [
        {
            "name": "one",
            "sheet_index": 0,
            "address": "A1",
            "scope_sheet_name": None,
            "scope_sheet_index": None,
            "book_scope": True,
        },
        {
            "name": "'Sheet 1'!two",
            "sheet_index": 0,
            "address": "C7:D8",
            "scope_sheet_name": "Sheet 1",
            "scope_sheet_index": 0,
            "book_scope": False,
        },
        {
            "name": "Sheet2!two",
            "sheet_index": 2,
            "address": "B3",
            "scope_sheet_name": "Sheet3",
            "scope_sheet_index": 2,
            "book_scope": False,
        },
        {
            "name": "two",
            "sheet_index": 1,
            "address": "A1:A2",
            "scope_sheet_name": None,
            "scope_sheet_index": None,
            "book_scope": True,
        },
    ],
    "sheets": [
        {
            "name": "Sheet 1",
            "values": [
                ["a", "b", "c", ""],
                [1.1, 2.2, 3.3, "2021-01-01T00:00:00.000Z"],
                [4.4, 5.5, 6.6, ""],
                ["", "", "", ""],
                ["", "", "", ""],
                ["", "", "", ""],
                ["", "", "", ""],
                ["", "", "", ""],
                ["", "", "", ""],
                ["Column1", "Column2", "", ""],
                [1.1, 2.2, "", ""],
                ["", "", "", ""],
                ["", "", "", ""],
                ["", "", "", ""],
                [1.1, 2.2, 3.3, ""],
                [4.4, 5.5, 6.6, ""],
                ["Total", "", 9.899999999999999, ""],
            ],
            "pictures": [
                {"name": "", "height": 13, "width": 35},
                {"name": "", "height": 40, "width": 62},
            ],
            "tables": [],
        },
        {
            "name": "Sheet2",
            "values": [["aa", "bb"], [11.1, 22.2]],
            "pictures": [],
            "tables": [],
        },
        {
            "name": "Sheet3",
            "values": [
                ["", "string"],
                [-1.1, 1.1],
                [True, False],
                ["2021-10-01T00:00:00.000Z", "2021-12-31T23:35:00.000Z"],
            ],
            "pictures": [],
            "tables": [],
        },
    ],
}
if __name__ == "__main__":
    import uvicorn

    uvicorn.run(
        "devserver:app",
        host="127.0.0.1",
        port=8000,
        reload=True,
        reload_dirs=[this_dir, this_dir.parent / "xlwings"],
        ssl_keyfile=this_dir / "localhost+2-key.pem",
        ssl_certfile=this_dir / "localhost+2.pem",
    )
