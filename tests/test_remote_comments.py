"""Threaded comments use queued mutations and asynchronous xlwings Lite reads."""

import datetime as dt
import sys
from types import ModuleType

import pytest

import xlwings as xw
from xlwings.pro import _xlremote as remote


@pytest.fixture
def anyio_backend():
    return "asyncio"


def book():
    data = {
        "client": "Office.js",
        "version": xw.__version__,
        "book": {"name": "B", "active_sheet_index": 0, "selection": "A1"},
        "names": [],
        "sheets": [
            {"name": "S1", "values": [[]], "pictures": [], "tables": []},
            {"name": "S2", "values": [[]], "pictures": [], "tables": []},
        ],
    }
    impl = remote.App(remote.Apps(), add_book=False).books.open(data)
    return xw.Book(impl=impl)


class Proxy:
    def __init__(self, value):
        self.value = value

    def to_py(self):
        return self.value


class Client:
    def __init__(self):
        self.calls = []

    async def getCommentAt(self, sheet, address):
        self.calls.append(("at", sheet, address))
        return Proxy({"id": "thread-1"}) if address == "$B$2" else None

    async def getComments(self, sheet=None):
        self.calls.append(("list", sheet))
        entries = [
            {"sheet": "S1", "address": "$B$2", "id": "thread-1"},
            {"sheet": "S2", "address": "$C$3", "id": "thread-2"},
        ]
        return Proxy(
            [entry for entry in entries if sheet is None or entry["sheet"] == sheet]
        )

    async def getCommentData(self, sheet, comment_id, address, key):
        self.calls.append(("read", sheet, comment_id, address, key))
        return {
            "text": "Review",
            "author": "Pat",
            "creation_date": "2026-09-24T10:00:00.000Z",
            "resolved": False,
            "location": Proxy({"sheet": "S2", "address": "$C$3"}),
        }[key]

    async def getCommentReplies(self, sheet, comment_id, address):
        self.calls.append(("replies", comment_id))
        return Proxy([{"id": "reply-1"}])

    async def getCommentReplyText(self, sheet, comment_id, address, reply_id):
        self.calls.append(("reply", reply_id))
        return "Done"


def install_client(monkeypatch):
    client = Client()
    js = ModuleType("js")
    js.xlwings = client
    monkeypatch.setitem(sys.modules, "js", js)
    monkeypatch.setattr(sys, "platform", "emscripten")
    return client


def test_comment_mutations_queue_without_rewriting_cells():
    wb = book()
    cell = wb.sheets[0]["B2"]
    comment = cell.add_comment("Review")
    comment.text = "Updated"
    comment.add_reply("Done")
    comment.resolve()
    comment.reopen()
    comment.delete()
    actions = wb.json()["actions"]
    assert [action["func"] for action in actions] == [
        "addComment",
        "setCommentText",
        "addCommentReply",
        "setCommentResolved",
        "setCommentResolved",
        "deleteComment",
    ]
    assert all(action["sheet_position"] == 0 for action in actions)
    assert actions[0]["args"] == ["$B$2", "Review"]
    assert actions[3]["args"] == [None, "$B$2", True]
    assert actions[4]["args"] == [None, "$B$2", False]
    assert all(action["func"] not in {"setValues", "setFormula"} for action in actions)


def test_comment_input_validation_and_sync_read_guidance():
    wb = book()
    cell = wb.sheets[0]["B2"]
    with pytest.raises(TypeError, match="string"):
        cell.add_comment(42)
    with pytest.raises(ValueError, match="empty"):
        cell.add_comment("")
    with pytest.raises(ValueError, match="single cell"):
        wb.sheets[0]["B2:C3"].add_comment("Review")
    assert wb.json()["actions"] == []
    with pytest.raises(NotImplementedError, match="get_comment"):
        cell.comment
    with pytest.raises(NotImplementedError, match="get_comments"):
        list(wb.sheets[0].comments)


@pytest.mark.anyio
async def test_comment_reads_and_reply_enumeration(monkeypatch):
    client = install_client(monkeypatch)
    wb = book()
    cell = wb.sheets[0]["B2"]
    assert await wb.sheets[0]["A1"].get_comment() is None
    comment = await cell.get_comment()
    assert isinstance(comment, xw.Comment)
    assert await comment.get_text() == "Review"
    assert await comment.get_author() == "Pat"
    assert await comment.get_creation_date() == dt.datetime(
        2026, 9, 24, 10, tzinfo=dt.timezone.utc
    )
    assert await comment.get_resolved() is False
    assert (await comment.get_location()).address == "$C$3"
    replies = await comment.get_replies()
    assert len(replies) == 1
    assert await replies[0].get_text() == "Done"
    assert ("read", "S1", "thread-1", "$B$2", "text") in client.calls


@pytest.mark.anyio
async def test_comment_collections_are_worksheet_and_workbook_scoped(monkeypatch):
    install_client(monkeypatch)
    wb = book()
    comments = await wb.get_comments()
    assert comments.count == 2
    assert comments[wb.sheets[0]["B2"]].impl.comment_id == "thread-1"
    sheet_comments = await wb.sheets[1].get_comments()
    assert sheet_comments.count == 1
    assert sheet_comments["C3"].impl.comment_id == "thread-2"
    with pytest.raises(KeyError):
        sheet_comments["A1"]
