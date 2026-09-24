"""Desktop bulk fill writes only visit cells selected by the matrix."""

import sys

import pytest

if sys.platform == "darwin":
    from xlwings import _xlmac as engine
elif sys.platform == "win32":
    from xlwings import _xlwindows as engine
else:
    engine = None


@pytest.mark.skipif(engine is None, reason="Requires a desktop xlwings engine")
def test_set_colors_preserves_unchanged_cells():
    writes = []

    class Cell:
        def __init__(self, row, column):
            self.row = row
            self.column = column

        @property
        def color(self):
            raise AssertionError("Bulk fill should not read existing colors")

        @color.setter
        def color(self, value):
            writes.append((self.row, self.column, value))

    class Target:
        def __call__(self, row, column):
            return Cell(row, column)

    engine.Range.set_colors(Target(), [[(0, 128, 0), ...], [None, (255, 255, 255)]])
    assert writes == [
        (1, 1, (0, 128, 0)),
        (2, 1, None),
        (2, 2, (255, 255, 255)),
    ]
