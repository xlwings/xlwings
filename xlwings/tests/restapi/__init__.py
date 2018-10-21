import os
import datetime as dt
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
        cls.app1 = xw.App()
        cls.app2 = xw.App()

        cls.wb1 = cls.app1.books.add()
        cls.wb1b = cls.app1.books.add()
        cls.wb2 = cls.app2.books.add()
        for wb in [cls.wb1, cls.wb2]:
            if len(wb.sheets) == 1:
                wb.sheets.add(after=1)

        sheet1 = cls.wb1.sheets[0]

        sheet1['A1'].value = [[1.1, 'a string'], [dt.datetime.now(), None]]
        sheet1['A1'].formula = '=1+1.1'
        chart = sheet1.charts.add()
        chart.set_source_data(sheet1['A1'])
        chart.chart_type = 'line'

        pic = os.path.abspath(os.path.join('..', 'sample_picture.png'))
        pic = sheet1.pictures.add(pic)
        pic.name = 'Picture 1'
        cls.wb1.sheets[0].range('B2:C3').name = 'Sheet1!myname1'
        cls.wb1.sheets[0].range('A1').name = 'myname2'
        cls.wb1.save('Book1.xlsx')
        cls.wb1 = xw.Book('Book1.xlsx')  # hack as save doesn't return the wb properly
        cls.app1.activate()

    @classmethod
    def tearDownClass(cls):
        wb_path = cls.wb1.fullname
        cls.app1.kill()
        cls.app2.kill()
        os.remove(wb_path)
