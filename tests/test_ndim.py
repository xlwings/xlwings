import pytest

import xlwings as xw


@pytest.fixture(scope="module")
def app():
    with xw.App(visible=False) as app:
        app.books.add()
        yield app


class TestNdimNatural:
    """Tests for ndim='natural' parameter"""

    def test_single_cell_scalar(self, app):
        """Single cell (1x1) returns scalar"""
        sheet = app.books[0].sheets[0]
        sheet["A1"].clear()
        sheet["A1"].value = 42
        result = sheet["A1"].options(ndim="natural").value
        assert result == 42

    def test_single_cell_string(self, app):
        """Single cell with string returns scalar string"""
        sheet = app.books[0].sheets[0]
        sheet["A1"].clear()
        sheet["A1"].value = "Hello"
        result = sheet["A1"].options(ndim="natural").value
        assert result == "Hello"

    def test_horizontal_range_1d(self, app):
        """Horizontal range (1xN) returns 1D array"""
        sheet = app.books[0].sheets[0]
        sheet["A1:D1"].clear()
        sheet["A1"].value = [1, 2, 3, 4]
        result = sheet["A1:D1"].options(ndim="natural").value
        assert result == [1.0, 2.0, 3.0, 4.0]

    def test_vertical_range_2d(self, app):
        """Vertical range (Nx1) returns 2D array"""
        sheet = app.books[0].sheets[0]
        sheet["A1:A3"].clear()
        sheet["A1"].value = [[1], [2], [3]]
        result = sheet["A1:A3"].options(ndim="natural").value
        assert result == [[1.0], [2.0], [3.0]]

    def test_2d_range(self, app):
        """2D range (NxM) returns 2D array"""
        sheet = app.books[0].sheets[0]
        sheet["A1:C2"].clear()
        sheet["A1"].value = [[1, 2, 3], [4, 5, 6]]
        result = sheet["A1:C2"].options(ndim="natural").value
        assert result == [[1.0, 2.0, 3.0], [4.0, 5.0, 6.0]]
