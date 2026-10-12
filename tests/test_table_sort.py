import sys
from types import SimpleNamespace
from unittest.mock import MagicMock, Mock

import pytest

from xlwings import main
from xlwings.pro import _xlremote


def table_stub(columns=3):
    implementation = Mock()
    selected = SimpleNamespace(
        range=SimpleNamespace(shape=(5, columns)),
        impl=implementation,
    )
    return selected, implementation


def test_sort_normalizes_keys_and_directions():
    selected, implementation = table_stub()
    main.Table.sort(selected, [3, 1], [False, True])
    implementation.sort.assert_called_once_with([3, 1], [False, True])

    selected, implementation = table_stub()
    main.Table.sort(selected, 2, ascending=False)
    implementation.sort.assert_called_once_with([2], [False])


@pytest.mark.parametrize(
    ("keys", "ascending", "error"),
    [
        ([], True, ValueError),
        (True, True, TypeError),
        ("Price", True, TypeError),
        ([1, 1], True, ValueError),
        ([4], True, ValueError),
        ([0], True, ValueError),
        ([1, 2], [True], ValueError),
        ([1], [1], TypeError),
    ],
)
def test_sort_rejects_invalid_arguments_before_dispatch(keys, ascending, error):
    selected, implementation = table_stub()
    with pytest.raises(error):
        main.Table.sort(selected, keys, ascending)
    implementation.sort.assert_not_called()


def test_remote_sort_queues_one_table_action():
    book = SimpleNamespace(append_json_action=Mock())
    sheet = SimpleNamespace(
        book=book,
        index=2,
        api={"tables": [{"name": "First"}, {"name": "Products"}]},
    )
    table = _xlremote.Table(sheet, 2)
    table.sort([3, 1], [False, True])
    book.append_json_action.assert_called_once_with(
        func="tableSort",
        args=[1, [3, 1], [False, True]],
        sheet_position=1,
    )


def test_mac_sort_uses_table_sort_fields():
    if sys.platform != "darwin":
        pytest.skip("macOS engine test")
    from xlwings import _xlmac

    table = _xlmac.Table.__new__(_xlmac.Table)
    table.xl = MagicMock()
    table.xl.show_headers.get.return_value = True

    table.sort([3, 1], [False, True])

    sort = table.xl.sort_object
    sort.sortfieldset.clear_sortfieldset.assert_called_once_with()
    calls = sort.sortfieldset.add_sortfield.call_args_list
    assert [call.kwargs["key"] for call in calls] == [
        table.xl.list_columns[3].range_object,
        table.xl.list_columns[1].range_object,
    ]
    assert [call.kwargs["order"] for call in calls] == [
        _xlmac.kw.sort_descending,
        _xlmac.kw.sort_ascending,
    ]
    sort.sort_orientation.set.assert_called_once_with(_xlmac.kw.sort_columns)
    sort.apply_sort.assert_called_once_with()


def test_mac_sort_without_headers_sorts_the_data_body(monkeypatch):
    if sys.platform != "darwin":
        pytest.skip("macOS engine test")
    from xlwings import _xlmac

    table = _xlmac.Table.__new__(_xlmac.Table)
    table.xl = MagicMock()
    table.xl.show_headers.get.return_value = False
    body = Mock()
    monkeypatch.setattr(_xlmac.Table, "data_body_range", property(lambda self: body))
    table.sort([2], [False])
    body.sort.assert_called_once_with([2], [False], False)
    table.xl.sort_object.apply_sort.assert_not_called()


def test_windows_sort_uses_table_sort_fields():
    if sys.platform != "win32":
        pytest.skip("Windows engine test")
    from xlwings import _xlwindows

    table = _xlwindows.Table(MagicMock())
    table.xl.ShowHeaders = True

    table.sort([3, 1], [False, True])

    sort = table.xl.Sort
    sort.SortFields.Clear.assert_called_once_with()
    assert sort.SortFields.Add.call_count == 2
    assert [call.kwargs["Order"] for call in sort.SortFields.Add.call_args_list] == [
        _xlwindows.constants.SortOrder.xlDescending,
        _xlwindows.constants.SortOrder.xlAscending,
    ]
    assert sort.Header == _xlwindows.constants.YesNoGuess.xlYes
    assert sort.Orientation == _xlwindows.constants.Constants.xlTopToBottom
    sort.Apply.assert_called_once_with()
