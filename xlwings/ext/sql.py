import sqlite3

from .. import arg, func, ret


def conv_value(value, col_is_str):
    if value is None:
        return "NULL"
    if col_is_str:
        return repr(str(value))
    elif isinstance(value, bool):
        return 1 if value else 0
    else:
        return repr(value)


@func
@arg("tables", expand="table", ndim=2)
@ret(expand="table")
def sql(query, *table_or_alias):
    return _sql(query, *table_or_alias)


@func
@arg("tables", expand="table", ndim=2)
def sql_dynamic(query, *table_or_alias):
    """Called if native dynamic arrays are available"""
    return _sql(query, *table_or_alias)


def _sql(query, *tables_or_aliases):
    """Excel formula: =SQL(query, ["alias1"], range1, ["alias2"], range2, ...)"""
    conn = sqlite3.connect(":memory:")
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
        types = [any(isinstance(row[j], str) for row in rows) for j in range(len(cols))]

        stmt = "CREATE TABLE %s (%s)" % (
            alias,
            ", ".join(
                "'%s' %s" % (col, "STRING" if typ else "REAL")
                for col, typ in zip(cols, types)
            ),
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
