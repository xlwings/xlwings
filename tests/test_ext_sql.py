"""
Tests for the in-Excel SQL extension (`xlwings.ext.sql`).

These call the Python functions directly with the nested lists that Excel would pass
for each range, so they don't need Excel.
"""

import datetime as dt
import sqlite3

import pytest

from xlwings.ext.sql import _sql, get_column_type, quote_identifier, sql, sql_dynamic

CUSTOMERS = [
    ["customer_id", "name"],
    [1.0, "Alice"],
    [2.0, "Bob"],
    [3.0, "Carol"],
]

ORDERS = [
    ["order_id", "customer_id", "amount"],
    [10.0, 1.0, 100.0],
    [11.0, 1.0, 50.0],
    [12.0, 2.0, 75.0],
]


# Aliases


def test_auto_aliases():
    assert _sql("SELECT name FROM a WHERE customer_id = 2", CUSTOMERS) == [
        ["name"],
        ["Bob"],
    ]
    assert _sql("SELECT COUNT(*) AS n FROM b", CUSTOMERS, ORDERS) == [["n"], [3]]


def test_auto_aliases_are_case_insensitive():
    assert _sql("SELECT COUNT(*) AS n FROM A", CUSTOMERS) == [["n"], [3]]


def test_explicit_aliases():
    res = _sql(
        "SELECT c.name, SUM(o.amount) AS total FROM customers c "
        "JOIN orders o ON c.customer_id = o.customer_id "
        "GROUP BY c.name ORDER BY c.name",
        [["customers"]],
        CUSTOMERS,
        [["orders"]],
        ORDERS,
    )
    assert res == [["name", "total"], ["Alice", 150.0], ["Bob", 75.0]]


def test_mixed_explicit_and_auto_aliases():
    # The auto alias counts all previous tables, including explicitly aliased ones
    res = _sql(
        "SELECT (SELECT COUNT(*) FROM customers) AS c, (SELECT COUNT(*) FROM b) AS o",
        [["customers"]],
        CUSTOMERS,
        ORDERS,
    )
    assert res == [["c", "o"], [3, 3]]


@pytest.mark.parametrize("alias", ["order", "my table", 'say "hi"', "select"])
def test_alias_needing_quotes(alias):
    query = "SELECT COUNT(*) AS n FROM %s" % quote_identifier(alias)
    assert _sql(query, [[alias]], CUSTOMERS) == [["n"], [3]]


def test_single_text_cell_is_an_alias_not_a_table():
    assert _sql("SELECT COUNT(*) AS n FROM hello", [["hello"]], CUSTOMERS) == [
        ["n"],
        [3],
    ]
    with pytest.raises(sqlite3.OperationalError):
        _sql("SELECT * FROM a", [["hello"]])


def test_single_number_cell_is_a_table():
    assert _sql("SELECT * FROM a", [[1.0]]) == [["1.0"]]


def test_trailing_alias_without_table_is_ignored():
    assert _sql("SELECT COUNT(*) AS n FROM a", CUSTOMERS, [["orphan"]]) == [
        ["n"],
        [3],
    ]
    with pytest.raises(sqlite3.OperationalError):
        _sql("SELECT * FROM orphan", CUSTOMERS, [["orphan"]])


def test_consecutive_aliases_use_the_last_one():
    args = [["first"]], [["second"]], CUSTOMERS
    assert _sql("SELECT COUNT(*) AS n FROM second", *args) == [["n"], [3]]
    with pytest.raises(sqlite3.OperationalError):
        _sql("SELECT * FROM first", *args)


# Column names


@pytest.mark.parametrize("col", ["it's", 'say "hi"', "my col", "order", "a-b"])
def test_column_name_needing_quotes(col):
    table = [[col], ["x"]]
    res = _sql("SELECT %s FROM a" % quote_identifier(col), table)
    assert res == [[col], ["x"]]


def test_numeric_column_name():
    res = _sql('SELECT "2024.0" FROM a', [[2024.0], [1.0]])
    assert res == [["2024.0"], [1.0]]


def test_duplicate_column_names_raise():
    with pytest.raises(sqlite3.OperationalError):
        _sql("SELECT * FROM a", [["x", "x"], [1.0, 2.0]])


# Values and types


def test_numbers_stay_numbers():
    res = _sql("SELECT amount, amount * 2 AS double FROM a", [["amount"], [1.5]])
    assert res == [["amount", "double"], [1.5, 3.0]]


def test_numeric_text_stays_text():
    table = [["code"], ["007"], ["1e3"], ["010"]]
    assert _sql("SELECT code FROM a", table) == [["code"], ["007"], ["1e3"], ["010"]]
    assert _sql("SELECT code FROM a WHERE code = '007'", table) == [
        ["code"],
        ["007"],
    ]


@pytest.mark.parametrize(
    "value",
    [
        """X"Y'Z""",
        "b\\'c",
        "it's",
        "back\\slash",
        "'; DROP TABLE a; --",
        "line\nbreak",
        "ünïcödé €",
    ],
)
def test_special_characters_in_values(value):
    table = [["v"], [value]]
    assert _sql("SELECT v FROM a", table) == [["v"], [value]]


def test_none_becomes_null():
    table = [["x", "y"], [1.0, None], [None, "b"]]
    assert _sql("SELECT * FROM a", table) == [["x", "y"], [1.0, None], [None, "b"]]
    assert _sql("SELECT COUNT(*) AS n FROM a WHERE x IS NULL", table) == [
        ["n"],
        [1],
    ]


def test_empty_string_stays_empty_string():
    table = [["x"], [""], ["a"]]
    assert _sql("SELECT x FROM a", table) == [["x"], [""], ["a"]]


def test_booleans():
    table = [["name", "active"], ["a", True], ["b", False], ["c", None]]
    assert _sql("SELECT * FROM a", table) == [
        ["name", "active"],
        ["a", True],
        ["b", False],
        ["c", None],
    ]
    assert _sql("SELECT name FROM a WHERE active", table) == [["name"], ["a"]]


def test_datetimes():
    table = [
        ["when"],
        [dt.datetime(2024, 1, 2, 3, 4, 5)],
        [None],
        [dt.datetime(2023, 12, 31)],
    ]
    assert _sql('SELECT * FROM a ORDER BY "when"', table) == [
        ["when"],
        [None],
        [dt.datetime(2023, 12, 31)],
        [dt.datetime(2024, 1, 2, 3, 4, 5)],
    ]


def test_datetime_comparison_with_iso_literal():
    table = [["when"], [dt.datetime(2024, 1, 2)], [dt.datetime(2023, 12, 31)]]
    res = _sql('SELECT "when" FROM a WHERE "when" >= \'2024-01-01\'', table)
    assert res == [["when"], [dt.datetime(2024, 1, 2)]]


def test_column_type_from_first_non_empty_value():
    table = [["x"], [None], [""], [2.0], [1.0]]
    # SQLite sorts NULL first, then numbers, then text
    assert _sql("SELECT x FROM a ORDER BY x", table) == [
        ["x"],
        [None],
        [1.0],
        [2.0],
        [""],
    ]


# Table shapes


def test_header_only_table():
    assert _sql("SELECT * FROM a", [["x", "y"]]) == [["x", "y"]]


def test_empty_column():
    table = [["x", "empty"], [1.0, None], [2.0, None]]
    assert _sql("SELECT * FROM a", table) == [
        ["x", "empty"],
        [1.0, None],
        [2.0, None],
    ]


def test_empty_result_returns_header():
    assert _sql("SELECT name FROM a WHERE 0", CUSTOMERS) == [["name"]]


def test_query_without_tables():
    assert _sql("SELECT 1 + 1 AS two") == [["two"], [2]]


# Errors


def test_invalid_query_raises():
    with pytest.raises(sqlite3.OperationalError):
        _sql("SELECT * FROM does_not_exist", CUSTOMERS)


def test_tables_are_not_shared_between_calls():
    _sql("SELECT * FROM a", CUSTOMERS)
    with pytest.raises(sqlite3.OperationalError):
        _sql("SELECT * FROM b", ORDERS)


# UDF entry points


@pytest.mark.parametrize("func", [sql, sql_dynamic])
def test_udf_entry_points(func):
    assert func("SELECT COUNT(*) AS n FROM a", CUSTOMERS) == [["n"], [3]]


# Helpers


@pytest.mark.parametrize(
    "values, expected",
    [
        ([1.0, 2.0], "REAL"),
        ([1, 2], "REAL"),
        (["a"], "TEXT"),
        ([True, False], "BOOLEAN"),
        ([dt.datetime(2024, 1, 1)], "DATETIME"),
        ([None, "", "a"], "TEXT"),
        ([None, "", 1.0], "REAL"),
        ([], "TEXT"),
        ([None, ""], "TEXT"),
    ],
)
def test_get_column_type(values, expected):
    assert get_column_type(values) == expected


@pytest.mark.parametrize(
    "name, expected",
    [
        ("a", '"a"'),
        ("my table", '"my table"'),
        ('say "hi"', '"say ""hi"""'),
        ("it's", '"it\'s"'),
        (2024.0, '"2024.0"'),
    ],
)
def test_quote_identifier(name, expected):
    assert quote_identifier(name) == expected
