import sys
from types import SimpleNamespace
from unittest.mock import MagicMock, Mock

import pytest

from xlwings import main
from xlwings.pro import _xlremote


def range_stub(shape=(4, 3), tables=()):
    implementation = Mock()
    selected = SimpleNamespace(
        shape=shape,
        row=2,
        column=3,
        sheet=SimpleNamespace(tables=tables),
        impl=implementation,
    )
    return selected, implementation


def test_sort_normalizes_keys_and_directions():
    selected, implementation = range_stub()
    main.Range.sort(selected, [2, 1], [False, True], has_headers=True)
    implementation.sort.assert_called_once_with([2, 1], [False, True], True)

    selected, implementation = range_stub()
    main.Range.sort(selected, 3, ascending=False)
    implementation.sort.assert_called_once_with([3], [False], False)


@pytest.mark.parametrize(
    ("keys", "ascending", "has_headers", "error"),
    [
        ([], True, False, ValueError),
        (True, True, False, TypeError),
        ([1, 1], True, False, ValueError),
        ([4], True, False, ValueError),
        ([0], True, False, ValueError),
        ([1, False], True, False, TypeError),
        ([1, 2], [True], False, ValueError),
        ([1], [1], False, TypeError),
        ([1], True, 1, TypeError),
    ],
)
def test_sort_rejects_invalid_arguments_before_dispatch(
    keys, ascending, has_headers, error
):
    selected, implementation = range_stub()
    with pytest.raises(error):
        main.Range.sort(selected, keys, ascending, has_headers)
    implementation.sort.assert_not_called()


def test_sort_rejects_header_only_range_and_table_intersection():
    selected, implementation = range_stub(shape=(1, 3))
    with pytest.raises(ValueError, match="at least two rows"):
        main.Range.sort(selected, 1, has_headers=True)
    implementation.sort.assert_not_called()

    table_range = SimpleNamespace(row=4, column=4, shape=(3, 2))
    selected, implementation = range_stub(tables=[SimpleNamespace(range=table_range)])
    with pytest.raises(ValueError, match="intersecting a table"):
        main.Range.sort(selected, 1)
    implementation.sort.assert_not_called()


def test_remote_sort_queues_one_range_action():
    book = SimpleNamespace(append_json_action=Mock())
    sheet = SimpleNamespace(book=book, index=2)
    selected = _xlremote.Range(sheet, (3, 4), (6, 7))
    selected.sort([2, 1], [False, True], True)
    book.append_json_action.assert_called_once_with(
        func="rangeSort",
        args=[[2, 1], [False, True], True],
        sheet_position=1,
        start_row=2,
        start_column=3,
        row_count=4,
        column_count=4,
    )


def test_mac_sort_uses_sort_fields_and_selected_range():
    if sys.platform != "darwin":
        pytest.skip("macOS engine test")
    from xlwings import _xlmac

    selected = _xlmac.Range.__new__(_xlmac.Range)
    selected._coords = (2, 3, 4, 3)
    selected.xl = Mock()
    sort = Mock()
    selected.sheet = SimpleNamespace(xl=MagicMock(sort_object=sort))

    selected.sort([2, 1], [False, True], True)

    assert sort.sortfieldset.add_sortfield.call_count == 2
    assert [
        call.kwargs["order"] for call in sort.sortfieldset.add_sortfield.call_args_list
    ] == [
        _xlmac.kw.sort_descending,
        _xlmac.kw.sort_ascending,
    ]
    sort.sortfieldset.clear_sortfieldset.assert_called_once_with()
    sort.set_sort_range.assert_called_once_with(rng=selected.xl)
    sort.sort_header.set.assert_called_once_with(_xlmac.kw.header_yes)
    sort.sort_orientation.set.assert_called_once_with(_xlmac.kw.sort_columns)
    sort.apply_sort.assert_called_once_with()


def test_windows_sort_uses_sort_fields_and_selected_range():
    if sys.platform != "win32":
        pytest.skip("Windows engine test")
    from xlwings import _xlwindows

    selected = _xlwindows.Range.__new__(_xlwindows.Range)
    selected._xl = MagicMock()
    selected._xl.Rows.Count = 4
    selected._xl.Columns.Count = 3
    selected._coords = (selected._xl.Worksheet, 2, 3, 4, 3)
    selected.sort([2, 1], [False, True], True)

    sort = selected._xl.Worksheet.Sort
    assert sort.SortFields.Add.call_count == 2
    sort.SortFields.Clear.assert_called_once_with()
    sort.SetRange.assert_called_once_with(selected._xl)
    assert sort.Header == _xlwindows.constants.YesNoGuess.xlYes
    assert sort.Orientation == _xlwindows.constants.Constants.xlTopToBottom
    sort.Apply.assert_called_once_with()
