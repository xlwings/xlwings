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
            # You can uncheck the header row in an Excel table
            header = (list(data.index.names) + list(data.columns)) if index else list(data.columns)
            # Replace None in the index with a unique number of spaces
            n_empty = len([i for i in header if i and ' ' in i])
            header = [f' ' * (i + n_empty + 1) if name is None else name for i, name in enumerate(header)]
            self.header_row_range.value = header
            self.range[1:, :].options(index=index, header=False).value = data
        else:
            # Without a table header, the table is deleted...
            self.show_headers = True
            self.range[1:, :].options(index=index, header=False).value = data
            self.show_headers = False
        return self
    else:
        raise TypeError(type_error_msg)
