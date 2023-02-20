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
def hello(data: dict = Body):
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
    assert book.name == "engines.xlsx", "engines.xlsx must be the active file"
    assert data == expected_body, "Body differs (Make sure to select cell Sheet1!A1)"
    book.app.alert("OK", title="Integration Test Read")
    return book.json()


@app.post("/integration-test-write")
def integration_test_write(data: dict = Body):
    book = xw.Book(json=data)
    assert (
        book.name == "integration_write.xlsx"
    ), "integration_write.xlsx must be the active file"
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
            "prompt": markupsafe.Markup(prompt.replace("\n", "<br>")),
            "title": title,
            "buttons": buttons,
            "mode": mode,
            "callback": callback,
        },
    )


@app.get("/xlwings/custom-functions-meta")
async def custom_functions_meta():
    return xw.pro.custom_functions_meta(custom_functions)


@app.get("/xlwings/custom-functions-code")
async def custom_functions_code():
    return PlainTextResponse(xw.pro.custom_functions_code(custom_functions))


@app.post("/xlwings/custom-functions-call")
async def custom_functions_call(data: dict = Body):
    rv = await xw.pro.custom_functions_call(data, custom_functions)
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


expected_body = {
    "client": "Office.js",
    "version": "dev",
    "book": {"name": "engines.xlsx", "active_sheet_index": 0, "selection": "A1"},
    "names": [
        {"name": "one", "sheet_index": 0, "address": "A1", "book_scope": True},
        {"name": "two", "sheet_index": 1, "address": "A1:A2", "book_scope": True},
        {"name": "two", "sheet_index": 0, "address": "C7:D8", "book_scope": False},
    ],
    "sheets": [
        {
            "name": "Sheet1",
            "values": [
                ["a", "b", "c", ""],
                [1, 2, 3, "2021-01-01T00:00:00.000Z"],
                [4, 5, 6, ""],
                ["", "", "", ""],
            ],
            "pictures": [],
        },
        {"name": "Sheet2", "values": [["aa", "bb"], [11, 22]], "pictures": []},
        {
            "name": "Sheet3",
            "values": [
                ["", "string"],
                [-1, 1],
                [True, False],
                # TODO: Custom datetime format not supported yet
                ["2021-10-01T00:00:00.000Z", 44561.9826388889],
            ],
            "pictures": [],
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
