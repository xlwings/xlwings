import asyncio
import sys
from types import SimpleNamespace
from unittest.mock import MagicMock, Mock

import pytest

from xlwings import main
from xlwings.pro import _xlremote


def public_range(shape=(4, 3), tables=()):
    impl = Mock()
    selected = SimpleNamespace(
        shape=shape,
        row=2,
        column=3,
        sheet=SimpleNamespace(tables=tables),
        impl=impl,
    )
    return selected, impl


def test_remove_duplicates_validates_and_dispatches_relative_columns():
    selected, impl = public_range()
    main.Range.remove_duplicates(selected, [2, 1], has_headers=True)
    impl.remove_duplicates.assert_called_once_with([2, 1], True)


@pytest.mark.parametrize(
    ("columns", "has_headers", "error"),
    [
        ([], False, ValueError),
        (True, False, TypeError),
        ([1, 1], False, ValueError),
        ([0], False, ValueError),
        ([4], False, ValueError),
        ([False], False, TypeError),
        ("1", False, TypeError),
        (1, 1, TypeError),
    ],
)
def test_remove_duplicates_rejects_bad_arguments_before_dispatch(
    columns, has_headers, error
):
    selected, impl = public_range()
    with pytest.raises(error):
        main.Range.remove_duplicates(selected, columns, has_headers)
    impl.remove_duplicates.assert_not_called()


def test_remove_duplicates_rejects_header_only_and_table_intersection():
    selected, impl = public_range(shape=(1, 3))
    with pytest.raises(ValueError, match="at least two rows"):
        main.Range.remove_duplicates(selected, 1, True)
    impl.remove_duplicates.assert_not_called()

    table_range = SimpleNamespace(row=3, column=4, shape=(2, 2))
    selected, impl = public_range(tables=[SimpleNamespace(range=table_range)])
    with pytest.raises(ValueError, match="intersecting a table"):
        main.Range.remove_duplicates(selected, 1)
    impl.remove_duplicates.assert_not_called()


@pytest.mark.parametrize(
    ("cell_type", "value_type", "error"),
    [
        (1, None, TypeError),
        ("other", None, ValueError),
        ("visible", "text", ValueError),
        ("constants", 1, TypeError),
        ("formulas", "other", ValueError),
    ],
)
def test_get_special_cells_rejects_invalid_arguments(cell_type, value_type, error):
    selected, impl = public_range()
    with pytest.raises(error):
        main.Range.get_special_cells(selected, cell_type, value_type)
    impl.get_special_cells.assert_not_called()


def test_get_special_cells_returns_public_areas_on_desktop_and_lite():
    selected, impl = public_range()
    areas = [Mock(), Mock()]
    impl.get_special_cells.return_value = areas
    result = main.Range.get_special_cells(selected, "visible")
    assert [area.impl for area in result] == areas
    impl.get_special_cells.assert_called_once_with("visible", None)

    async def read_areas(*args):
        return areas

    impl.get_special_cells = read_areas
    result = asyncio.run(main.Range.get_special_cells(selected, "constants", "numbers"))
    assert [area.impl for area in result] == areas


def test_remote_remove_duplicates_queues_one_action():
    book = SimpleNamespace(append_json_action=Mock())
    sheet = SimpleNamespace(book=book, index=2)
    selected = _xlremote.Range(sheet, (3, 4), (6, 7))
    selected.remove_duplicates([2, 1], True)
    book.append_json_action.assert_called_once_with(
        func="rangeRemoveDuplicates",
        args=[[2, 1], True],
        sheet_position=1,
        start_row=2,
        start_column=3,
        row_count=4,
        column_count=4,
    )


def test_remote_get_special_cells_reads_host(monkeypatch):
    async def read(*args):
        return SimpleNamespace(to_py=lambda: ["$D$3", "$F$5:$G$6"])

    host = SimpleNamespace(getSpecialCells=Mock(side_effect=read))
    monkeypatch.setitem(sys.modules, "js", SimpleNamespace(xlwings=host))
    monkeypatch.setattr(_xlremote.sys, "platform", "emscripten")
    sheet = SimpleNamespace(
        name="Data",
        index=1,
        book=SimpleNamespace(append_json_action=Mock()),
    )
    selected = _xlremote.Range(sheet, (3, 4), (6, 7))
    result = asyncio.run(selected.get_special_cells("visible", None))
    assert [area.address for area in result] == ["$D$3", "$F$5:$G$6"]
    host.getSpecialCells.assert_called_once_with("Data", "$D$3:$G$6", "visible", None)


def test_mac_remove_duplicates_is_explicitly_unavailable():
    if sys.platform != "darwin":
        pytest.skip("macOS engine test")
    from xlwings import _xlmac

    selected = _xlmac.Range.__new__(_xlmac.Range)
    with pytest.raises(NotImplementedError, match="AppleScript"):
        selected.remove_duplicates([1], False)


def test_mac_special_cells_uses_native_command():
    if sys.platform != "darwin":
        pytest.skip("macOS engine test")
    from xlwings import _xlmac

    selected = _xlmac.Range.__new__(_xlmac.Range)
    selected.xl = MagicMock()
    selected._coords = (2, 3, 4, 3)
    selected.sheet = SimpleNamespace(xl=MagicMock())
    selected.xl.special_cells.return_value.get_address.return_value = "$A$1,$C$3:$D$4"
    areas = selected.get_special_cells("constants", "numbers")
    assert len(areas) == 2
    selected.xl.special_cells.assert_called_once_with(
        type=_xlmac.kw.cell_type_constants, value=1
    )

    selected._coords = (2, 3, 1, 1)
    selected.sheet.book = SimpleNamespace(app=SimpleNamespace(xl=MagicMock()))
    selected.sheet.book.app.xl.intersect.return_value = _xlmac.kw.missing_value
    assert selected.get_special_cells("formulas", None) == []
    selected.sheet.book.app.xl.intersect.return_value = MagicMock()
    assert selected.get_special_cells("visible", None) == [selected]


def test_windows_remove_duplicates_and_special_cells_are_native():
    if sys.platform != "win32":
        pytest.skip("Windows engine test")
    from xlwings import _xlwindows

    selected = _xlwindows.Range.__new__(_xlwindows.Range)
    selected._xl = MagicMock()
    selected.remove_duplicates([2, 1], True)
    selected._xl.RemoveDuplicates.assert_called_once_with(
        Columns=[2, 1], Header=_xlwindows.constants.YesNoGuess.xlYes
    )
    selected._xl.SpecialCells.return_value.Areas = [MagicMock(), MagicMock()]
    selected._xl.Application.Intersect.return_value.Areas = [MagicMock(), MagicMock()]
    areas = selected.get_special_cells("visible", None)
    assert len(areas) == 2
    selected._xl.Application.Intersect.assert_called_once_with(
        selected._xl.SpecialCells.return_value, selected._xl
    )
    selected._xl.SpecialCells.assert_called_once_with(
        Type=_xlwindows.constants.CellType.xlCellTypeVisible
    )
    selected._xl.Application.Intersect.return_value = None
    assert selected.get_special_cells("visible", None) == []
