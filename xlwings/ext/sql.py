from .. import func, arg, ret, serve
import sqlite3


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
@arg("tables", expand='table', ndim=2)
@ret(expand='table')
def sql(query, *tables):
    conn = sqlite3.connect(':memory:')

    c = conn.cursor()

    for i, table in enumerate(tables):
        cols = table[0]
        rows = table[1:]
        types = [
            any(type(row[j]) is str for row in rows)
            for j in range(len(cols))
        ]
        name = chr(65 + i)

        stmt = "CREATE TABLE %s (%s)" % (
            name,
            ", ".join("'%s' %s" % (col, "STRING" if typ else "REAL") for col, typ in zip(cols, types))
        )
        c.execute(stmt)

        if rows:
            stmt = "INSERT INTO %s VALUES %s" % (
                name,
                ", ".join(
                    "(%s)" % ", ".join(
                        conv_value(value, type)
                        for value, typ in zip(row, types)
                    )
                    for row in rows
                )
            )
            c.execute(stmt)

    res = []
    c.execute(query)
    res.append([x[0] for x in c.description])
    for row in c:
        res.append(list(row))

    return res

if __name__ == "__main__":
    serve()
