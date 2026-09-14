from types import SimpleNamespace

import pytest

from xlwings.pro import _xlcalamine


@pytest.mark.parametrize(
    ("bounds", "expected_address"),
    [
        (((0, 0), (16, 3)), "$A$1:$D$17"),
        (((4, 2), (9, 3)), "$C$5:$D$10"),
        (((1, 1), (1, 1)), "$B$2"),
        (None, "$A$1"),
    ],
)
def test_used_range_converts_calamine_bounds(monkeypatch, bounds, expected_address):
    calls = []

    def fake_get_used_range(path, sheet_index):
        calls.append((path, sheet_index))
        return bounds

    monkeypatch.setattr(
        _xlcalamine.xlwingslib, "get_used_range", fake_get_used_range, raising=False
    )
    sheet = _xlcalamine.Sheet(book=SimpleNamespace(fullname="book.xlsx"), sheet_index=2)

    assert sheet.used_range.address == expected_address
    assert calls == [("book.xlsx", 1)]
