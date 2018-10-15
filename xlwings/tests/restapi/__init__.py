import unittest

import xlwings as xw
from xlwings.rest.api import api


class TestCase(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        # Flask
        api.config['TESTING'] = True
        cls.client = api.test_client()

        # xlwings
        cls.app1 = xw.App(visible=False)
        cls.app2 = xw.App(visible=False)

        cls.wb1 = cls.app1.books.add()
        cls.wb2 = cls.app2.books.add()
        for wb in [cls.wb1, cls.wb2]:
            if len(wb.sheets) == 1:
                wb.sheets.add(after=1)
                wb.sheets.add(after=2)
                wb.sheets[0].select()

    @classmethod
    def tearDownClass(cls):
        cls.app1.kill()
        cls.app2.kill()
