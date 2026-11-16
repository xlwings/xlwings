import datetime as dt
import sqlite3

from .. import arg, func, ret


@func
@arg("arg", expand="table", ndim=2)
@ret(expand="table")
def sql(query, *arg):
    # Singular arg to make is show up correctly in Excel
    return _sql(query, *arg)


@func
@arg("arg", expand="table", ndim=2)
def sql_dynamic(query, *arg):
    """Called if native dynamic arrays are available"""
    return _sql(query, *arg)


def conv_value(value):
    if isinstance(value, dt.datetime):
        # Same format as SQLite's datetime(), so that comparisons with it work
        return value.isoformat(" ")
    return value


def quote_identifier(name):
    return '"%s"' % str(name).replace('"', '""')


def get_column_type(column_values):
    types = set()
    for val in column_values:
        if val in (None, ""):
            continue
        elif isinstance(val, dt.datetime):
            types.add("DATETIME")
        elif isinstance(val, bool):
            types.add("BOOLEAN")
        elif isinstance(val, str):
            types.add("TEXT")
        else:
            types.add("REAL")
    if not types:
        return "TEXT"
    if len(types) == 1:
        return types.pop()
    # Mixed columns, e.g., "n/a" and numbers: BLOB affinity stores values as they are
    return "BLOB"


def _sql(query, *tables_or_aliases):
    """Excel formula: =SQL(query, ["alias1"], range1, ["alias2"], range2, ...)"""

    def convert_datetime(bytestring):
        return dt.datetime.fromisoformat(bytestring.decode("utf-8"))

    def convert_boolean(value):
        return value.decode("utf-8").lower() in ("1", "true")

    sqlite3.register_converter("DATETIME", convert_datetime)
    sqlite3.register_converter("BOOLEAN", convert_boolean)

    conn = sqlite3.connect(":memory:", detect_types=sqlite3.PARSE_DECLTYPES)
    c = conn.cursor()

    tables = []
    current_alias = None

    # Process arguments into (alias, table) pairs
    for table_or_alias in tables_or_aliases:
        if (
            len(table_or_alias) == 1
            and len(table_or_alias[0]) == 1
            and isinstance(table_or_alias[0][0], str)
        ):
            current_alias = table_or_alias[0][0]
        else:
            if current_alias is None:
                # Auto-assign alias (A, B, C...)
                current_alias = chr(65 + len(tables))
            tables.append((current_alias, table_or_alias))
            current_alias = None

    for alias, table in tables:
        cols = table[0]
        rows = table[1:]
        types = []
        for j in range(len(cols)):
            column_values = (row[j] for row in rows)
            types.append(get_column_type(column_values))

        stmt = "CREATE TABLE %s (%s)" % (
            quote_identifier(alias),
            ", ".join(
                "%s %s" % (quote_identifier(col), typ) for col, typ in zip(cols, types)
            ),
        )
        c.execute(stmt)

        if rows:
            c.executemany(
                "INSERT INTO %s VALUES (%s)"
                % (quote_identifier(alias), ", ".join("?" * len(cols))),
                ([conv_value(value) for value in row] for row in rows),
            )

    res = []
    c.execute(query)
    res.append([x[0] for x in c.description])
    for row in c:
        res.append(list(row))

    return res
