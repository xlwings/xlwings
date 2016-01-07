from .. import xlplatform

# Optional imports
try:
    import numpy as np
except ImportError:
    np = None

try:
    import pandas as pd
except ImportError:
    pd = None

class Converter(object):
    def read(self, rng):
        return xlplatform.clean_xl_data(xlplatform.get_value_from_range(rng.xl_range))

    def write(self, rng, value):
        if isinstance(value, (tuple, list)):
            if len(value) == 0:
                return
            if isinstance(value[0], (tuple, list)):
                row2 = rng.row1 + len(value) - 1
                col2 = rng.col1 + len(value[0]) - 1
            else:
                row2 = rng.row1 + len(value) - 1
                col2 = rng.col1
                value = [[x] for x in value]
            value = xlplatform.prepare_xl_data(value)
        else:
            row2 = rng.row2
            col2 = rng.col2
            value = xlplatform.prepare_xl_data([[value]])[0][0]

        xlplatform.set_value(xlplatform.get_range_from_indices(rng.xl_sheet, rng.row1, rng.col1, row2, col2), value)


class DefaultConverter(Converter):
    def __init__(self, ndim=None):
        if ndim not in (None, 1, 2):
            raise ValueError("'ndim' argument must be None, 1 or 2")
        self.ndim = ndim

    def read(self, rng):
        value = super(DefaultConverter, self).read(rng)
        if self.ndim is None:
            return value
        elif self.ndim == 1:
            if isinstance(value, (list, tuple)):
                if len(value) > 0 and isinstance(value[0], (list, tuple)):
                    if len(value[0]) > 0:
                        return [x[0] for x in value]
                    else:
                        return []
                else:
                    return value
            else:
                return [value]
        elif self.ndim == 2:
            if isinstance(value, (list, tuple)):
                if len(value) > 0 and isinstance(value[0], (list, tuple)):
                    value
                else:
                    return [value]
            else:
                return [[value]]


default = DefaultConverter()

def ndim(n):
    return DefaultConverter(ndim=n)