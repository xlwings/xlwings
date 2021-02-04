try:
    import pandas as pd
except ImportError:
    pd = None


def update(self, data):
    type_error_msg = 'Currently, only pandas DataFrames are supported by update'
    if pd:
        if not isinstance(data, pd.DataFrame):
            raise TypeError(type_error_msg)
        col_diff = len(self.range.columns) - len(data.columns) - len(data.index.names)
        nrows = len(self.data_body_range.rows) if self.data_body_range else 1
        row_diff = nrows - len(data.index)
        if col_diff > 0:
            self.range[:, len(self.range.columns) - col_diff:].delete()
        if row_diff > 0 and self.data_body_range:
            self.data_body_range[len(self.data_body_range.rows) - row_diff:, :].delete()
        self.header_row_range.value = list(data.index.names) + list(data.columns)
        self.range[1:, :].options(index=True, header=False).value = data
        return self
    else:
        raise TypeError(type_error_msg)
