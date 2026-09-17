import asyncio
from unittest import mock

import pytest

import xlwings as xw
from xlwings import XlwingsError
from xlwings.pro import _xlremote as R


def _book():
    payload = {
        "client": "Office.js",
        "version": xw.__version__,
        "book": {"name": "B", "active_sheet_index": 0, "selection": "A1"},
        "names": [],
        "sheets": [
            {
                "name": "Sheet1",
                "values": [[1, 2], [3, 4]],
                "pictures": [],
                "tables": [],
            }
        ],
    }
    impl = R.App(R.Apps(), add_book=False).books.open(payload)
    return xw.Book(impl=impl)


def _entries():
    return [
        {"type": "CellValue", "stop_if_true": True},
        {"type": "DataBar", "stop_if_true": None},
        {"type": "PresetCriteria", "stop_if_true": False},
    ]


def test_public_types_are_exported():
    assert xw.ConditionalFormat.__name__ == "ConditionalFormat"
    assert xw.ConditionalFormats.__name__ == "ConditionalFormats"


def test_unloaded_collection_can_clear_without_a_read():
    book = _book()
    formats = book.sheets[0]["A1:B2"].conditional_formats

    formats.clear()

    action = book.impl.json()["actions"][-1]
    assert action["func"] == "clearConditionalFormats"
    assert action["args"] == []
    assert (action["row_count"], action["column_count"]) == (2, 2)


def test_unloaded_collection_points_inspection_at_async_getter():
    formats = _book().sheets[0]["A1"].conditional_formats
    with pytest.raises(NotImplementedError, match=r"get_conditional_formats\(\)"):
        len(formats)


def test_async_getter_is_lite_only_off_emscripten():
    with pytest.raises(NotImplementedError, match="only supported in xlwings Lite"):
        asyncio.run(_book().sheets[0]["A1"].get_conditional_formats())


def test_async_getter_preserves_order_and_unknown_rules():
    rng = _book().sheets[0]["A1:B2"]

    async def fake(self, key, method=None):
        assert key == "conditional_formats"
        return _entries()

    with mock.patch.object(R.Range, "_get_range_data", fake):
        formats = asyncio.run(rng.get_conditional_formats())

    assert len(formats) == 3
    assert [rule.type for rule in formats] == ["cell_value", "data_bar", "unknown"]
    assert [rule.stop_if_true for rule in formats] == [True, None, False]
    assert formats[0].type == "cell_value"
    assert formats(2).type == "data_bar"


def test_multiple_deletes_keep_snapshot_positions_aligned():
    book = _book()
    rng = book.sheets[0]["A1:B2"]

    async def fake(self, key, method=None):
        return _entries()

    with mock.patch.object(R.Range, "_get_range_data", fake):
        formats = asyncio.run(rng.get_conditional_formats())

    rules = list(formats)
    rules[0].delete()
    rules[1].delete()

    actions = book.impl.json()["actions"]
    assert [action["func"] for action in actions] == [
        "deleteConditionalFormat",
        "deleteConditionalFormat",
    ]
    assert [action["args"][0] for action in actions] == [0, 0]
    assert [rule.type for rule in formats] == ["unknown"]


def test_deleted_rule_cannot_be_deleted_twice():
    rng = _book().sheets[0]["A1"]

    async def fake(self, key, method=None):
        return _entries()[:1]

    with mock.patch.object(R.Range, "_get_range_data", fake):
        rule = asyncio.run(rng.get_conditional_formats())[0]
    rule.delete()
    with pytest.raises(XlwingsError, match="no longer in its collection"):
        rule.delete()


def test_clear_updates_a_loaded_snapshot():
    rng = _book().sheets[0]["A1"]

    async def fake(self, key, method=None):
        return _entries()

    with mock.patch.object(R.Range, "_get_range_data", fake):
        formats = asyncio.run(rng.get_conditional_formats())
    formats.clear()
    assert len(formats) == 0
