try:
    import pandas as pd
except ImportError:
    pd = None


def update(self, data, index):
    type_error_msg = 'Currently, only pandas DataFrames are supported by update'
    if pd:
        if not isinstance(data, pd.DataFrame):
            raise TypeError(type_error_msg)
        col_diff = len(self.range.columns) - len(data.columns) - (len(data.index.names) if index else 0)
        nrows = len(self.data_body_range.rows) if self.data_body_range else 1
        row_diff = nrows - len(data.index)
        if col_diff > 0:
            self.range[:, len(self.range.columns) - col_diff:].delete()
        if row_diff > 0 and self.data_body_range:
            self.data_body_range[len(self.data_body_range.rows) - row_diff:, :].delete()
        if self.header_row_range:
            # Tables with 'Header Row' checked
            header = (list(data.index.names) + list(data.columns)) if index else list(data.columns)
            # Replace None in the index with a unique number of spaces
            n_empty = len([i for i in header if i and ' ' in i])
            header = [f' ' * (i + n_empty + 1) if name is None else name for i, name in enumerate(header)]
            self.header_row_range.value = header
            self.range[1, 0].options(index=index, header=False).value = data
        else:
            # Tables with 'Header Row' unchecked
            self.range[0, 0].options(index=index, header=False).value = data
            # If the top-left cell isn't empty, it doesn't manage to resize the columns automatically
            data_rows = len(data)
            data_cols = len(data.columns) if not index else len(data.columns) + len(data.index.names)
            self.resize(self.range[0, 0].resize(row_size=data_rows,
                                                column_size=data_cols))
        return self
    else:
        raise TypeError(type_error_msg)
