import pytest

import xlwings as xw
from xlwings.pro import _xlremote as remote


def remote_table():
    app = remote.App(remote.Apps(), add_book=False)
    book = app.books.open(
        {
            "client": "Office.js",
            "version": xw.__version__,
            "book": {"name": "B", "active_sheet_index": 0},
            "sheets": [
                {
                    "name": "S",
                    "values": [[]],
                    "tables": [
                        {
                            "name": "T",
                            "range_address": "B2:C4",
                            "row_count": 3,
                            "column_count": 2,
                            "show_headers": True,
                            "show_totals": False,
                        }
                    ],
                }
            ],
        }
    )
    return xw.Book(impl=book).sheets[0].tables[0]


def test_remote_table_rows_queue_and_local_count():
    table = remote_table()
    rows = table.rows
    assert len(rows) == 2
    assert [row.index for row in rows] == [1, 2]
    assert rows[-1].index == 2
    appended = rows.add(["new", 3])
    inserted = rows.add(index=1)
    assert appended.index == 3
    assert inserted.index == 1
    assert len(rows) == 4
    rows[1].delete()
    assert len(rows) == 3
    actions = table.parent.book.json()["actions"]
    assert [(action["func"], action["args"]) for action in actions] == [
        ("addTableRow", [0, 2, ["new", 3]]),
        ("addTableRow", [0, 0, None]),
        ("deleteTableRow", [0, 1]),
    ]


@pytest.mark.parametrize(
    ("values", "index", "error"),
    [
        ([1], None, ValueError),
        ([[1, 2]], None, ValueError),
        ([1, object()], None, TypeError),
        ([1, float("nan")], None, ValueError),
        ([1, 2], True, TypeError),
        ([1, 2], 0, IndexError),
        ([1, 2], 4, IndexError),
    ],
)
def test_remote_table_row_validation_before_queue(values, index, error):
    table = remote_table()
    with pytest.raises(error):
        table.rows.add(values, index=index)
    assert table.parent.book.json()["actions"] == []


def test_remote_row_sync_range_points_to_async_getter():
    row = remote_table().rows[0]
    with pytest.raises(NotImplementedError, match="await row.get_range"):
        row.range
