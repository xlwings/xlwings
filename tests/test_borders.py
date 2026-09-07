"""Round-trip tests for Range.borders against a real Excel (Windows and macOS).

Some expectations encode what Excel actually does rather than what was asked
for. They were measured on Excel for Mac on 2026-09-07; re-measure before
assuming other hosts or versions behave the same.
"""

import sys
from pathlib import Path

import pytest

import xlwings as xw

this_dir = Path(__file__).resolve().parent

OUTSIDE = ["edge_top", "edge_bottom", "edge_left", "edge_right"]
INSIDE = ["inside_vertical", "inside_horizontal"]
DIAGONALS = ["diagonal_down", "diagonal_up"]
GRID_SIDES = OUTSIDE + INSIDE
ALL_SIDES = GRID_SIDES + DIAGONALS
LINE_STYLES = [
    "continuous",
    "dash",
    "dash_dot",
    "dash_dot_dot",
    "dot",
    "double",
    "slant_dash_dot",
]
WEIGHTS = ["hairline", "thin", "medium", "thick"]


@pytest.fixture(scope="module")
def app():
    with xw.App(visible=False) as app:
        app.books.open(this_dir / "test book.xlsx")
        yield app


@pytest.fixture
def rng(app):
    # A multi-cell range, so that the inside borders exist
    rng = app.books[0].sheets[0]["B2:D4"]
    rng.borders.clear()
    yield rng
    rng.borders.clear()


def test_api(rng):
    assert rng.borders.api is not None
    assert rng.borders["edge_top"].api is not None


def test_iteration(rng):
    assert len(rng.borders) == 8
    assert [border.line_style for border in rng.borders] == ["none"] * 8


def test_no_borders_by_default(rng):
    for side in ALL_SIDES:
        assert rng.borders[side].line_style == "none"
    assert rng.borders.line_style == "none"


@pytest.mark.parametrize("side", ALL_SIDES)
@pytest.mark.parametrize("line_style", LINE_STYLES)
def test_line_style_round_trip(rng, side, line_style):
    rng.borders[side].line_style = line_style
    assert rng.borders[side].line_style == line_style


@pytest.mark.parametrize("side", ALL_SIDES)
@pytest.mark.parametrize("weight", WEIGHTS)
def test_weight_round_trip(rng, side, weight):
    rng.borders[side].line_style = "continuous"
    rng.borders[side].weight = weight
    assert rng.borders[side].weight == weight
    assert rng.borders[side].line_style == "continuous"


@pytest.mark.parametrize("side", ALL_SIDES)
def test_color_round_trip(rng, side):
    rng.borders[side].line_style = "continuous"
    rng.borders[side].color = (255, 0, 0)
    assert rng.borders[side].color == (255, 0, 0)
    rng.borders[side].color = "#00ff00"
    assert rng.borders[side].color == (0, 255, 0)
    rng.borders[side].color = 0xFF0000  # Excel's BGR integer: blue
    assert rng.borders[side].color == (0, 0, 255)


def test_enums_are_accepted(rng):
    rng.borders[xw.BorderIndex.edge_left].line_style = xw.BorderLineStyle.dash
    rng.borders[xw.BorderIndex.edge_left].weight = xw.BorderWeight.medium
    assert rng.borders["edge_left"].line_style == "dash"
    assert rng.borders["edge_left"].weight == "medium"
    rng.borders[xw.BorderIndex.edge_left].line_style = xw.BorderLineStyle.none
    assert rng.borders["edge_left"].line_style == "none"


@pytest.mark.parametrize("value", [None, "none"])
def test_line_style_none_removes_border(rng, value):
    rng.borders["edge_top"].line_style = "continuous"
    rng.borders["edge_top"].line_style = value
    assert rng.borders["edge_top"].line_style == "none"


@pytest.mark.skipif(sys.platform != "darwin", reason="measured on macOS only")
def test_removed_border_has_no_color_on_mac(rng):
    rng.borders["edge_top"].line_style = "continuous"
    rng.borders["edge_top"].color = (255, 0, 0)
    rng.borders["edge_top"].line_style = None
    assert rng.borders["edge_top"].color is None


def test_collection_properties_target_grid_sides_only(rng):
    rng.borders.line_style = "continuous"
    for side in GRID_SIDES:
        assert rng.borders[side].line_style == "continuous"
    for side in DIAGONALS:
        assert rng.borders[side].line_style == "none"
    assert rng.borders.line_style == "continuous"

    rng.borders.weight = "medium"
    for side in GRID_SIDES:
        assert rng.borders[side].weight == "medium"
    assert rng.borders.weight == "medium"

    rng.borders.color = "#0000ff"
    for side in GRID_SIDES:
        assert rng.borders[side].color == (0, 0, 255)
    assert rng.borders.color == (0, 0, 255)


def test_collection_getters_none_when_grid_sides_differ(rng):
    rng.borders.set("all", line_style="continuous", weight="thin", color=(0, 0, 0))
    assert rng.borders.line_style == "continuous"
    assert rng.borders.weight == "thin"
    assert rng.borders.color == (0, 0, 0)

    rng.borders["edge_top"].line_style = "double"
    assert rng.borders.line_style is None
    rng.borders["edge_top"].line_style = "continuous"
    assert rng.borders.line_style == "continuous"

    rng.borders["inside_horizontal"].weight = "thick"
    assert rng.borders.weight is None

    rng.borders["edge_left"].color = (255, 0, 0)
    assert rng.borders.color is None


def test_collection_getters_ignore_diagonals(rng):
    rng.borders.set("all", line_style="continuous", weight="thin")
    rng.borders["diagonal_up"].line_style = "double"
    rng.borders["diagonal_down"].weight = "thick"
    assert rng.borders.line_style == "continuous"
    assert rng.borders.weight == "thin"


@pytest.mark.parametrize("address", ["B2", "B2:D2", "B2:B4", "B2:D4"])
def test_collection_getters_ignore_nonexistent_inside_borders(rng, address):
    borders = rng.sheet[address].borders
    borders.set(line_style="continuous", weight="medium", color="#ff0000")
    assert borders.line_style == "continuous"
    assert borders.weight == "medium"
    assert borders.color == (255, 0, 0)
    borders["edge_top"].color = "#00ff00"
    assert borders.color is None
    borders.clear()
    assert borders.line_style == "none"
    assert borders.color is None


@pytest.mark.parametrize(
    "side,cell,cell_side",
    [
        ("edge_top", "C2", "edge_top"),
        ("edge_bottom", "C4", "edge_bottom"),
        ("edge_left", "B3", "edge_left"),
        ("edge_right", "D3", "edge_right"),
        ("inside_vertical", "C3", "edge_right"),
        ("inside_horizontal", "C3", "edge_bottom"),
        ("diagonal_down", "C3", "diagonal_down"),
        ("diagonal_up", "C3", "diagonal_up"),
    ],
)
@pytest.mark.parametrize(
    "attribute,value,initial",
    [
        ("line_style", "double", "continuous"),
        ("line_style", None, "continuous"),
        ("weight", "thick", "thin"),
        ("color", "#00ff00", (255, 0, 0)),
    ],
)
def test_border_getters_none_when_segments_differ(
    rng, side, cell, cell_side, attribute, value, initial
):
    borders = rng.borders
    borders.set("everything", line_style="continuous", weight="thin", color="#ff0000")
    border = borders[side]
    assert getattr(border, attribute) == initial
    setattr(rng.sheet[cell].borders[cell_side], attribute, value)
    assert getattr(border, attribute) is None
    if side in GRID_SIDES:
        assert getattr(borders, attribute) is None
    # Reuse the same wrapper after restoring uniform formatting.
    borders.set(side, line_style="continuous", weight="thin", color="#ff0000")
    assert getattr(border, attribute) == initial


def test_edge_getters_ignore_other_cells(rng):
    rng.borders.set("edge_top", line_style="continuous", weight="thin", color="#ff0000")
    rng.sheet["C3"].borders.set(
        "edge_top", line_style="double", weight="thick", color="#00ff00"
    )
    border = rng.borders["edge_top"]
    assert border.line_style == "continuous"
    assert border.weight == "thin"
    assert border.color == (255, 0, 0)


def test_set_groups(rng):
    rng.borders.set("outside", line_style="double", color=(255, 0, 0))
    for side in OUTSIDE:
        assert rng.borders[side].line_style == "double"
        assert rng.borders[side].color == (255, 0, 0)
    for side in INSIDE + DIAGONALS:
        assert rng.borders[side].line_style == "none"

    # "medium dashed" is one of Excel's built-in border styles, so both
    # attributes survive. Something like dot + hairline wouldn't: Excel
    # would keep the style and reset the weight to thin.
    rng.borders.set("inside", line_style="dash", weight="medium")
    for side in INSIDE:
        assert rng.borders[side].line_style == "dash"
        assert rng.borders[side].weight == "medium"
    for side in OUTSIDE:
        assert rng.borders[side].line_style == "double"

    rng.borders.set(["edge_top", "diagonal_up"], line_style="dash")
    assert rng.borders["edge_top"].line_style == "dash"
    assert rng.borders["diagonal_up"].line_style == "dash"
    assert rng.borders["edge_bottom"].line_style == "double"


def test_set_everything(rng):
    rng.borders.set("everything", line_style="continuous")
    for side in ALL_SIDES:
        assert rng.borders[side].line_style == "continuous"


def test_clear_scopes(rng):
    rng.borders.set("everything", line_style="continuous")
    rng.borders.clear("inside")
    for side in INSIDE:
        assert rng.borders[side].line_style == "none"
    for side in OUTSIDE + DIAGONALS:
        assert rng.borders[side].line_style == "continuous"

    rng.borders.clear("all")
    for side in GRID_SIDES:
        assert rng.borders[side].line_style == "none"
    for side in DIAGONALS:
        assert rng.borders[side].line_style == "continuous"

    rng.borders.clear()
    for side in ALL_SIDES:
        assert rng.borders[side].line_style == "none"


def test_removal_wins_when_weight_and_color_are_supplied(rng):
    rng.borders.set("outside", line_style="continuous")
    rng.borders.set("outside", weight="thick", color=(255, 0, 0), line_style=None)
    for side in OUTSIDE:
        assert rng.borders[side].line_style == "none"


def test_incompatible_combination_keeps_the_line_style(rng):
    # dash_dot_dot + thick isn't representable in Excel. With the documented
    # colour, weight, line-style order, the style is what survives.
    rng.borders.set("edge_top", line_style="dash_dot_dot", weight="thick")
    assert rng.borders["edge_top"].line_style == "dash_dot_dot"
    if sys.platform == "darwin":
        # Measured 2026-09-07: Excel for Mac downgrades the weight to thin
        assert rng.borders["edge_top"].weight == "thin"
    else:
        assert rng.borders["edge_top"].weight != "thick"
