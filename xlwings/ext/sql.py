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


def conv_value(value, col_is_str):
    if value is None:
        return "NULL"
    if col_is_str:
        return repr(str(value))
    elif isinstance(value, dt.datetime):
        return value.isoformat()
    else:
        return repr(value)


def get_column_type(column_values):
    for val in column_values:
        if val in (None, ""):
            continue
        elif isinstance(val, dt.datetime):
            return "DATETIME"
        elif isinstance(val, bool):
            return "BOOLEAN"
        elif isinstance(val, str):
            return "STRING"
        return "REAL"


def _sql(query, *tables_or_aliases):
    """Excel formula: =SQL(query, ["alias1"], range1, ["alias2"], range2, ...)"""

    def convert_datetime(bytestring):
        return dt.datetime.fromisoformat(bytestring.decode("utf-8"))

    def convert_boolean(value):
        print(value)
        return True if value.decode("utf-8").lower() == "true" else False

    sqlite3.register_converter("DATETIME", convert_datetime)
    sqlite3.register_converter("BOOLEAN", convert_boolean)

    conn = sqlite3.connect(":memory:", detect_types=sqlite3.PARSE_DECLTYPES)
    c = conn.cursor()

    tables = []
    current_alias = None

    # Process arguments into (alias, table) pairs
    for table_or_alias in tables_or_aliases:
        if len(table_or_alias[0]) == 1 and isinstance(table_or_alias[0][0], str):
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
            alias,
            ", ".join("'%s' %s" % (col, typ) for col, typ in zip(cols, types)),
        )
        c.execute(stmt)

        if rows:
            stmt = "INSERT INTO %s VALUES %s" % (
                alias,
                ", ".join(
                    "(%s)"
                    % ", ".join(
                        conv_value(value, type) for value, typ in zip(row, types)
                    )
                    for row in rows
                ),
            )
            # Fixes values like these:
            # sql('SELECT a FROM a', [['a', 'b'], ["""X"Y'Z""", 'd']])
            stmt = stmt.replace("\\'", "''")
            c.execute(stmt)

    res = []
    c.execute(query)
    res.append([x[0] for x in c.description])
    for row in c:
        res.append(list(row))

    return res
