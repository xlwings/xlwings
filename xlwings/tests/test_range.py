import unittest

from win32com.client.dynamic import Dispatch
from xlwings import Range, Workbook


class TestDemonstrateRangeBug(unittest.TestCase):
    """ This test case demonstrates a bug where a Range cannot be created with
    a named range referring to a sheet other than the active sheet.
    """
    def setUp(self):
        self.excel_application = Dispatch('Excel.Application')
        self.excel_application.Visible = False
        self.excel_application.DisplayAlerts = False
        self.workbook = self.excel_application.WorkBooks.Add()

        # use win32com method to add a name referring to a cell on sheet2
        self.sheet1 = self.excel_application.ActiveSheet
        self.sheet2 = self.workbook.Sheets.Add()
        cell_ref = "$A$1"
        ext_ref_str = "={}!{}".format(self.sheet2.Name, cell_ref)
        self.ext_named_range = "external_reference"
        self.workbook.Names.Add(Name=self.ext_named_range,
                                RefersTo=ext_ref_str)

        # convert the COM workbook to an xlwings.Workbook
        self.work_book = Workbook(xl_workbook=self.workbook)

    def test_create_range_on_sheet_referenced(self):
        self.sheet2.Activate()
        range_from_sheet_2 = Range(self.ext_named_range)
        self.assertIsNone(range_from_sheet_2.value)

    def test_create_range_on_different_sheet(self):
        """This test demonstrates the bug.
        """
        self.sheet1.Activate()
        range_from_sheet_1 = Range(self.ext_named_range)
        self.assertIsNone(range_from_sheet_1.value)

    def tearDown(self):
        self.excel_application.Quit()
