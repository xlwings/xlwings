import os
import inspect
import unittest
import types

import xlwings as xw

this_dir = os.path.abspath(os.path.dirname(inspect.getfile(inspect.currentframe())))

SPEC = None  # This was used to support Excel 2011: '/Applications/Microsoft Office 2011/Microsoft Excel'


class TestBase(unittest.TestCase):
    def __init__(self, methodName):
        super(TestBase, self).__init__(methodName)

        # Patch the test method being run to skip the test if it
        # throws NotImplementedError. This allows us to not consider
        # such tests failures, though they will still show up (as
        # skipped tests).
        test_method = getattr(self, methodName)

        def wrapped_method(self, *args, **kwargs):
            try:
                return test_method(*args, **kwargs)
            except NotImplementedError:
                self.skipTest("Test body threw NotImplementedError.")

        setattr(self, methodName, types.MethodType(wrapped_method, self))

    @classmethod
    def setUpClass(cls):
        cls.app1 = xw.App(visible=False, spec=SPEC)
        cls.app2 = xw.App(visible=False, spec=SPEC)

    def setUp(self):
        self.wb1 = self.app1.books.add()
        self.wb2 = self.app2.books.add()
        for wb in [self.wb1, self.wb2]:
            if len(wb.sheets) == 1:
                wb.sheets.add(after=1)
                wb.sheets.add(after=2)
                wb.sheets[0].select()

    # def tearDown(self):
    #     for app in [self.app1, self.app2]:
    #         app.books[-1].close()
    #
    # @classmethod
    # def tearDownClass(cls):
    #     cls.app1.kill()
    #     cls.app2.kill()
