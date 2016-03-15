import xlwings as xw
from xlwings.conversion import Accessor
from xlwings.conversion.standard import ValueAccessor
import numbers


class OneAccessor(Accessor):

    writes_types = int

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
        return ValueAccessor.reader(options).append_stage(cls.AddStage(options=options))

    @classmethod
    def writer(cls, options):
        return ValueAccessor.writer(options).insert_stage(cls.SubtractStage(options=options), index=1)

    @classmethod
    def router(cls, value, rng, options):
        if isinstance(value, (int, float, list, tuple)):
            return cls
        else:
            return super(OneAccessor, cls).router(value, rng, options)

OneAccessor.install_for(int)

wb = xw.Workbook.active()
xw.Range('A20').options(subtract_one=True).value = 1
