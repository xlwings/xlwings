from fastapi import Body, FastAPI, status
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import PlainTextResponse

import xlwings as xw

app = FastAPI()


@app.post("/hello")
def hello(data: dict = Body):
    # Instantiate a Book object with the deserialized request body
    with xw.Book(json=data) as book:
        # Use xlwings as usual
        sheet = book.sheets[0]
        cell = sheet["A1"]
        if cell.value == "Hello xlwings!":
            cell.value = "Bye xlwings!"
        else:
            cell.value = "Hello xlwings!"

        # Pass the following back as the response
        return book.json()


@app.exception_handler(Exception)
async def exception_handler(request, exception):
    # This handles all exceptions, so you may want to make this more restrictive
    return PlainTextResponse(
        str(exception), status_code=status.HTTP_500_INTERNAL_SERVER_ERROR
    )


# Office Scripts and custom functions in Excel on the web require CORS
cors_app = CORSMiddleware(
    app=app,
    allow_origins="*",
    allow_methods=["POST"],
    allow_headers=["*"],
    allow_credentials=True,
)


if __name__ == "__main__":
    import uvicorn

    uvicorn.run("main:cors_app", host="127.0.0.1", port=8000, reload=True)
