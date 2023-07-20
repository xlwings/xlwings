from pathlib import Path

import pytest

import xlwings as xw

this_dir = Path(__file__).resolve().parent


def test_col_name_raise_exception():
    """Assert that an ``IndexError`` is raised when a column value is given not in range 1 , 16384"""
    # Excel Columns start at 1
    with pytest.raises(IndexError) as exception:
        xw.utils.col_name(0)
    assert (
        str(exception.value)
        == 'Invalid column index "0". Column index needs to be between 1 and 16384'
    )
    # Excel Columns end at 16384
    with pytest.raises(IndexError) as exception:
        xw.utils.col_name(16385)
    assert (
        str(exception.value)
        == 'Invalid column index "16385". Column index needs to be between 1 and 16384'
    )


def test_col_name_no_errors_for_all_valid_columns():
    """No Exceptions should be raised in range 1, 16384"""
    for x in range(1, 16385):
        xw.utils.col_name(x)


def test_col_name_correct_letters():
    """assert that the results for a few samples are as expected"""

    assert xw.utils.col_name(1) == "A"
    assert xw.utils.col_name(45) == "AS"
    assert xw.utils.col_name(77) == "BY"
    assert xw.utils.col_name(16384) == "XFD"
