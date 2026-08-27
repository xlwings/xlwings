"""
Tests for the async Range.get_formula() on the remote (Office.js) backend.

Formulas deliberately aren't part of the payload the client sends with every
request (that would grow every request for data most scripts never read), so
the sync `formula` getter raises and formulas are fetched on demand instead.

`js` isn't available in the test environment and get_formula() gates on
``sys.platform == "emscripten"``, so the fixture below fakes just enough of
that surface to run the real logic.
"""

import sys
from types import ModuleType, SimpleNamespace

import pytest

import xlwings as xw
from xlwings.pro import _xlremote as R


def _book_json():
    return {
        "client": "Office.js",
        "version": xw.__version__,
        "book": {"name": "B", "active_sheet_index": 0, "selection": "A1"},
        "names": [],
        "sheets": [
            {
                "name": "S1",
                "values": [[None] * 5] * 5,
                "pictures": [],
                "tables": [],
            }
        ],
    }


@pytest.fixture
def book():
    impl = R.App(R.Apps(), add_book=False).books.open(_book_json())
    return xw.Book(impl=impl)


@pytest.fixture
def fake_js(monkeypatch):
    """Fake sys.platform + js so get_formula() runs its real logic.

    getRangeFormulas records the (sheet_name, address) it was called with and
    returns a formula grid matching the requested shape.
    """
    monkeypatch.setattr(sys, "platform", "emscripten")

    calls = []

    async def get_range_formulas(sheet_name, address):
        calls.append((sheet_name, address))
        tuple1, tuple2 = xw.utils.a1_to_tuples(address)
        if tuple2:
            nrows = tuple2[0] - tuple1[0] + 1
            ncols = tuple2[1] - tuple1[1] + 1
        else:
            nrows = ncols = 1
        grid = [
            [f"=R{tuple1[0] + r}C{tuple1[1] + c}" for c in range(ncols)]
            for r in range(nrows)
        ]
        return SimpleNamespace(to_py=lambda: grid)

    js = ModuleType("js")
    js.xlwings = SimpleNamespace(getRangeFormulas=get_range_formulas)
    monkeypatch.setitem(sys.modules, "js", js)
    return calls


@pytest.fixture
def anyio_backend():
    return "asyncio"


def _shape_of(obj):
    """Nesting shape of a read result: scalar, flat list, or nested list."""
    if not isinstance(obj, list):
        return "scalar"
    return "nested" if obj and isinstance(obj[0], list) else "flat"


# --- sync getter still raises, pointing at the async API ---


def test_formula_getter_raises_with_a_hint(book):
    with pytest.raises(NotImplementedError, match="get_formula"):
        book.sheets[0]["A1"].formula


# --- async getter ---


@pytest.mark.anyio
async def test_get_formula_single_cell_returns_a_scalar(book, fake_js):
    assert await book.sheets[0]["A1"].get_formula() == "=R1C1"


@pytest.mark.anyio
async def test_get_formula_horizontal_range_returns_a_flat_list(book, fake_js):
    # Squeezed like `.value`, not nested
    assert await book.sheets[0]["A1:C1"].get_formula() == [
        "=R1C1",
        "=R1C2",
        "=R1C3",
    ]


@pytest.mark.anyio
async def test_get_formula_vertical_range_returns_a_flat_list(book, fake_js):
    assert await book.sheets[0]["A1:A3"].get_formula() == [
        "=R1C1",
        "=R2C1",
        "=R3C1",
    ]


@pytest.mark.anyio
async def test_get_formula_2d_range_returns_a_nested_list(book, fake_js):
    assert await book.sheets[0]["A1:B2"].get_formula() == [
        ["=R1C1", "=R1C2"],
        ["=R2C1", "=R2C2"],
    ]


@pytest.mark.anyio
async def test_get_formula_shape_matches_value_reads(book, fake_js):
    # The whole point: get_formula() squeezes exactly like `.value` does
    sheet = book.sheets[0]
    sheet["A1:C1"].value = [1, 2, 3]
    sheet["E1:E3"].value = [[1], [2], [3]]
    for address in ["A1", "A1:C1", "E1:E3", "A1:B2"]:
        value_shape = _shape_of(sheet[address].value)
        formula_shape = _shape_of(await sheet[address].get_formula())
        assert value_shape == formula_shape, address


@pytest.mark.anyio
@pytest.mark.parametrize(
    "address, expected",
    [
        ("A1", [["=R1C1"]]),
        ("A1:C1", [["=R1C1", "=R1C2", "=R1C3"]]),
        ("A1:A3", [["=R1C1"], ["=R2C1"], ["=R3C1"]]),
    ],
)
async def test_get_formula_respects_ndim_2(book, fake_js, address, expected):
    assert await book.sheets[0][address].options(ndim=2).get_formula() == expected


@pytest.mark.anyio
async def test_get_formula_respects_ndim_1_on_a_single_cell(book, fake_js):
    assert await book.sheets[0]["A1"].options(ndim=1).get_formula() == ["=R1C1"]


@pytest.mark.anyio
async def test_get_formula_passes_sheet_name_and_address(book, fake_js):
    await book.sheets[0]["B2:C3"].get_formula()
    assert fake_js == [("S1", "$B$2:$C$3")]


@pytest.mark.anyio
async def test_get_formula_uses_the_ranges_own_sheet(book, fake_js):
    book.sheets[0].name = "Renamed"
    await book.sheets[0]["A1"].get_formula()
    assert fake_js[0][0] == "Renamed"


@pytest.mark.anyio
async def test_get_formula_works_on_a_lazy_book(book, fake_js):
    # Formulas are fetched on demand, so they don't need loaded values
    impl = R.App(R.Apps(), add_book=False).books.open(_book_json(), lazy=True)
    lazy_book = xw.Book(impl=impl)
    assert await lazy_book.sheets[0]["A1"].get_formula() == "=R1C1"


# --- platform gate ---


@pytest.mark.anyio
async def test_get_formula_raises_off_emscripten(book):
    with pytest.raises(NotImplementedError, match="xlwings Lite"):
        await book.sheets[0]["A1"].get_formula()
