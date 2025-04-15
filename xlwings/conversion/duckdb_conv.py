import csv
import os
import tempfile

from ..utils import xlserial_to_datetime

try:
    import duckdb
except ImportError:
    duckdb = None


class DuckRelation:
    def __init__(self, rel, con, name, temp_file):
        self.rel = rel
        self.con = con
        self.name = name
        self._temp_file = temp_file

    def __getattr__(self, name):
        return getattr(self.rel, name)

    def close(self):
        self.con.close()
        if os.path.exists(self._temp_file):
            os.remove(self._temp_file)


if duckdb:
    from . import Converter, Options

    def _parse_dates(duck_relation, parse_dates):
        if isinstance(parse_dates, (str, int)):
            parse_dates = [parse_dates]

        cols_to_parse_names = set()
        all_columns = duck_relation.columns
        for item in parse_dates:
            if isinstance(item, int):
                if 0 <= item < len(all_columns):
                    cols_to_parse_names.add(all_columns[item])
                else:
                    raise IndexError(f"Column index {item} is out of bounds.")
            elif isinstance(item, str):
                cols_to_parse_names.add(item)
            else:
                raise TypeError(
                    "Items in 'parse_dates' must be (a list of) column names (str) or indices (int)."
                )

        duck_relation.con.create_function(
            "xlserial_to_datetime", xlserial_to_datetime, return_type="TIMESTAMP"
        )
        cols = [
            f'xlserial_to_datetime("{col}") AS "{col}"'
            if col in cols_to_parse_names
            else f'"{col}"'
            for col in all_columns
        ]
        query = f"SELECT {', '.join(cols)} FROM {duck_relation.name}"
        return DuckRelation(
            duck_relation.con.sql(query),
            duck_relation.con,
            duck_relation.name,
            duck_relation._temp_file,
        )

    class DuckdbConverter(Converter):
        @classmethod
        def base_reader(cls, options):
            return super(DuckdbConverter, cls).base_reader(
                Options(options).override(ndim=2)
            )

        @classmethod
        def read_value(cls, value, options):
            parse_dates = options.get("parse_dates")
            name = options.get("name", "rel")

            with tempfile.NamedTemporaryFile(
                mode="w", delete=False, suffix=".csv", newline=""
            ) as tmp:
                writer = csv.writer(
                    tmp, delimiter=",", quotechar='"', quoting=csv.QUOTE_MINIMAL
                )
                writer.writerows(value)

            con = duckdb.connect()
            rel = con.read_csv(tmp.name, sep=",", header=True, quotechar='"')
            con.register(name, rel)
            drel = DuckRelation(rel, con, name, tmp.name)
            if parse_dates is not None:
                drel = _parse_dates(drel, parse_dates)
            return drel

        @classmethod
        def write_value(cls, value, options):
            rel = value
            result = [rel.columns]
            result.extend([list(row) for row in rel.fetchall()])
            return result

    DuckdbConverter.register("duckdb", "DuckDB", DuckRelation, duckdb.DuckDBPyRelation)
