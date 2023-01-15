import datetime as dt

import jinja2
import markupsafe
from dateutil import tz
from fastapi import Body, FastAPI, Request
from fastapi.responses import HTMLResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates

import xlwings as xw

app = FastAPI()


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
                ["2021-10-01T00:00:00.000Z", "2021-12-31T23:35:00.000Z"],
            ],
            "pictures": [],
        },
    ],
}


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
    # Click "Integration Test" on the Taskpane
    # with tests/test_engines/engines.xlsx open
    # NOTE: Select "A1" on Sheet1
    book = xw.Book(json=data)
    assert book.name == "engines.xlsx"
    import json

    print(json.dumps(data))
    assert data == expected_body
    book.app.alert("OK", title="Integration Test Read")
    return book.json()


@app.post("/integration-test-write")
def integration_test_write(data: dict = Body):
    book = xw.Book(json=data)
    assert book.name == "integration_write.xlsx"
    sheet1 = book.sheets[0]
    sheet1["B2"].value = [
        [None, "string"],
        [-1, 1],
        [-1.1, 1.1],
        [True, False],
        [
            dt.date(2021, 10, 1),
            dt.datetime(2021, 12, 31, 23, 35, 12, tzinfo=tz.gettz("Europe/Paris")),
        ],
    ]
    return book.json()


@app.get("/alert", response_class=HTMLResponse)
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


app.mount("/static", StaticFiles(directory="static"), name="static")
app.mount("/", StaticFiles(directory="build"), name="home")
StaticFiles.is_not_modified = lambda *args, **kwargs: False  # Never cache static files

# Add the xlwings alert template as source
loader = jinja2.ChoiceLoader(
    [
        jinja2.FileSystemLoader("build"),
        jinja2.PackageLoader("xlwings", "html"),
    ]
)
templates = Jinja2Templates(directory="build", loader=loader)


if __name__ == "__main__":
    import uvicorn

    uvicorn.run(
        "server:app",
        host="127.0.0.1",
        port=8000,
        reload=True,
        reload_dirs=["."],
        reload_includes=["*.py", "*.html", "*.js", "*.css"],
        ssl_keyfile="localhost+2-key.pem",
        ssl_certfile="localhost+2.pem",
    )
