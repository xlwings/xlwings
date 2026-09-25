import asyncio
import sys
from types import SimpleNamespace
from unittest.mock import MagicMock, Mock

import pytest

from xlwings import main
from xlwings.pro import _xlremote


def public_range(implementation=None):
    implementation = implementation or Mock()
    return main.Range(impl=implementation), implementation


def test_find_returns_desktop_result_without_await():
    selected, implementation = public_range()
    found = Mock()
    implementation.find.return_value = found

    result = selected.find(
        "needle", whole=True, direction="backward", order="columns", match_case=True
    )

    assert isinstance(result, main.Range)
    assert result.impl is found
    implementation.find.assert_called_once_with(
        "needle", True, "backward", "columns", True
    )
    implementation.find.return_value = None
    assert selected.find("missing") is None


def test_find_returns_awaitable_result_for_lite():
    selected, implementation = public_range()
    found = Mock()

    async def find(*args):
        return found

    implementation.find = find
    pending = selected.find("needle")
    result = asyncio.run(pending)
    assert isinstance(result, main.Range)
    assert result.impl is found


@pytest.mark.parametrize(
    ("args", "kwargs", "error"),
    [
        (("",), {}, ValueError),
        ((1,), {}, TypeError),
        (("x",), {"whole": 1}, TypeError),
        (("x",), {"match_case": 1}, TypeError),
        (("x",), {"direction": "up"}, ValueError),
        (("x",), {"order": "diagonal"}, ValueError),
    ],
)
def test_find_validates_before_dispatch(args, kwargs, error):
    selected, implementation = public_range()
    with pytest.raises(error):
        selected.find(*args, **kwargs)
    implementation.find.assert_not_called()


def test_replace_all_queues_without_returning_count():
    selected, implementation = public_range()
    selected.replace_all("old", "", whole=True, match_case=True)
    implementation.replace_all.assert_called_once_with("old", "", True, True)


@pytest.mark.parametrize(
    ("old", "new", "kwargs", "error"),
    [
        ("", "new", {}, ValueError),
        (1, "new", {}, TypeError),
        ("old", None, {}, TypeError),
        ("old", "new", {"whole": 1}, TypeError),
        ("old", "new", {"match_case": 1}, TypeError),
    ],
)
def test_replace_all_validates_before_dispatch(old, new, kwargs, error):
    selected, implementation = public_range()
    with pytest.raises(error):
        selected.replace_all(old, new, **kwargs)
    implementation.replace_all.assert_not_called()


def test_remote_replace_all_emits_one_scoped_action():
    book = SimpleNamespace(append_json_action=Mock())
    sheet = SimpleNamespace(book=book, index=2)
    selected = _xlremote.Range(sheet, (3, 4), (6, 7))
    selected.replace_all("old", "new", True, False)
    book.append_json_action.assert_called_once_with(
        func="rangeReplaceAll",
        args=["old", "new", True, False],
        sheet_position=1,
        start_row=2,
        start_column=3,
        row_count=4,
        column_count=4,
    )


def test_remote_find_reads_host_and_returns_range(monkeypatch):
    async def find_range(*args):
        return "$E$5"

    host = SimpleNamespace(findRange=Mock(side_effect=find_range))
    monkeypatch.setitem(sys.modules, "js", SimpleNamespace(xlwings=host))
    monkeypatch.setattr(_xlremote.sys, "platform", "emscripten")
    sheet = SimpleNamespace(
        name="Data",
        index=1,
        api={"names": [], "tables": []},
        book=SimpleNamespace(append_json_action=Mock()),
    )
    selected = _xlremote.Range(sheet, (3, 4), (6, 7))

    result = asyncio.run(selected.find("needle", True, "backward", "columns", True))

    assert isinstance(result, _xlremote.Range)
    assert result.address == "$E$5"
    host.findRange.assert_called_once_with(
        "Data", "$D$3:$G$6", "needle", True, "backward", "columns", True
    )

    async def no_match(*args):
        return None

    host.findRange.side_effect = no_match
    assert asyncio.run(selected.find("absent", False, "forward", "rows", False)) is None


def test_mac_find_and_replace_use_native_commands():
    if sys.platform != "darwin":
        pytest.skip("macOS engine test")
    from xlwings import _xlmac

    selected = _xlmac.Range.__new__(_xlmac.Range)
    selected._coords = (2, 3, 4, 3)
    selected.xl = MagicMock()
    app = SimpleNamespace(display_alerts=True)
    selected.sheet = SimpleNamespace(xl=MagicMock(), book=SimpleNamespace(app=app))
    selected.xl.find.return_value.get_address.return_value = "$E$5"

    result = selected.find("needle", True, "forward", "columns", True)

    assert isinstance(result, _xlmac.Range)
    assert selected.xl.find.call_args.kwargs["after_"] == (
        selected.sheet.xl.rows[5].columns[5]
    )
    assert selected.xl.find.call_args.kwargs["search_order"] == _xlmac.kw.by_columns
    assert selected.xl.find.call_args.kwargs["look_at"] == _xlmac.kw.whole
    selected.replace_all("old", "new", False, True)
    selected.xl.replace.assert_called_once_with(
        what="old",
        replacement="new",
        look_at=_xlmac.kw.part,
        search_order=_xlmac.kw.by_rows,
        match_case=True,
        match_byte=False,
    )
    assert app.display_alerts is True

    selected.xl.find.return_value = _xlmac.kw.missing_value
    assert selected.find("absent", False, "forward", "rows", False) is None


def test_windows_find_and_replace_use_native_commands():
    if sys.platform != "win32":
        pytest.skip("Windows engine test")
    from xlwings import _xlwindows

    selected = _xlwindows.Range.__new__(_xlwindows.Range)
    selected._xl = MagicMock()
    selected._xl.Rows.Count = 4
    selected._xl.Columns.Count = 3
    selected._coords = (selected._xl.Worksheet, 2, 3, 4, 3)
    found = MagicMock()
    selected._xl.Find.return_value = found

    result = selected.find("needle", True, "backward", "columns", True)

    assert isinstance(result, _xlwindows.Range)
    assert result.xl is found
    kwargs = selected._xl.Find.call_args.kwargs
    assert kwargs["After"] is selected._xl.Cells(1, 1)
    assert kwargs["LookIn"] == _xlwindows.constants.FindLookIn.xlValues
    assert kwargs["LookAt"] == _xlwindows.constants.LookAt.xlWhole
    assert kwargs["SearchOrder"] == _xlwindows.constants.SearchOrder.xlByColumns
    assert kwargs["SearchDirection"] == (
        _xlwindows.constants.SearchDirection.xlPrevious
    )
    selected.replace_all("old", "new", False, False)
    assert selected._xl.Replace.call_args.kwargs["LookAt"] == (
        _xlwindows.constants.LookAt.xlPart
    )
    assert selected._xl.Replace.call_args.kwargs["Replacement"] == "new"
