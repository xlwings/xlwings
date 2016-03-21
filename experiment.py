import xlwings as xw
from xlwings.conversion import Accessor, ConverterAccessor, ValueAccessor, CleanDataFromReadStage, PandasDataFrameConverter
import numbers

import numpy as np
import pandas as pd


class OneAccessor(Accessor):

    class AddStage(object):
        def __init__(self, options):
            self.options = options

        def __call__(self, ctx):
            if self.options.get('add_one', False):
                ctx.value = [[cell + 1 if isinstance(cell, numbers.Number) else cell for cell in row] for row in ctx.value]

    class SubtractStage(object):
        def __init__(self, options):
            self.options = options

        def __call__(self, ctx):
            if self.options.get('subtract_one', False):
                ctx.value = [[cell - 1 if isinstance(cell, numbers.Number) else cell for cell in row] for row in ctx.value]

    @classmethod
    def reader(cls, options):
        return ValueAccessor.reader(options).insert_stage(cls.AddStage(options=options), after=CleanDataFromReadStage)

    @classmethod
    def writer(cls, options):
        return ValueAccessor.writer(options).insert_stage(cls.SubtractStage(options=options), index=1)


OneAccessor.register(float)

wb = xw.Workbook.active()
xw.Range('A20').value = None
xw.Range('A20').options(subtract_one=True).value = 1.0


class DataFrameDropna(ConverterAccessor):

    base = PandasDataFrameConverter

    @classmethod
    def read_value(cls, df, options):
        dropna = options.get('dropna', True)
        if dropna:
            return df.dropna()
        else:
            return df

    @classmethod
    def write_value(cls, df, options):
        dropna = options.get('dropna', True)
        if dropna:
            df = df.dropna()
        return df

DataFrameDropna.register(pd.DataFrame)  # RecursionError

wb = xw.Workbook.active()
df = pd.DataFrame([[1,10],[2,np.nan], [3, 30]])
xw.Range('H1').options(DataFrameDropna).value = df
