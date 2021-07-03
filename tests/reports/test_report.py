import os
import sys
import unittest

import numpy as np
import pandas as pd
from pandas.testing import assert_frame_equal
from numpy.testing import assert_array_equal
from matplotlib.figure import Figure

import xlwings as xw
from xlwings.pro.reports import create_report, Image
from xlwings.pro import Markdown

# Test data
text1 = """\
# Title

Text **bold** and *italic*

* a first {{ mybullet }}
* a second bullet

# Another title

this has a line break
new line
"""

fig = Figure(figsize=(4, 3))
ax = fig.add_subplot(111)
ax.plot([1, 2, 3, 4, 5])

df1 = pd.DataFrame(index=['r0', 'r1'], columns=['c0', 'c1'], data=[[1., 1.], [1., 1.]])
df2 = pd.DataFrame({'name': ['a', 'b', 'c', 'd', 'e'],
                    'b': [4, 2, 6, 6, 9],
                    'c': [1, 2, 5, 7, 8],
                    'd': [1, 1, 1, 6, 7]})
data = dict(mystring='stringtest', myfloat=12.12, substring='substringtest',
            df1=df1, df2=df2, mydict={'df': df1}, pic=Image(os.path.abspath('xlwings.jpg')),
            fig=fig, markdown_cell=Markdown(text1), markdown_shape=Markdown(text1), mybullet='bullet')


class TestCreateReport(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.wb = create_report('template1.xlsx', 'output.xlsx', **data)

    @classmethod
    def tearDownClass(cls):
        xw.Book('output.xlsx').app.quit()

    def test_string(self):
        self.assertEqual(self.wb.sheets[0]['A1'].value, 'stringtest')

    def test_float(self):
        self.assertAlmostEqual(self.wb.sheets[0]['B1'].value, 12.12)

    def test_substring(self):
        self.assertAlmostEqual(self.wb.sheets[0]['C1'].value, 'This is text with a substringtest.')

    def test_df(self):
        assert_frame_equal(self.wb.sheets[0]['A2'].options(pd.DataFrame, expand='table').value,
                           data['df1'])

    def test_df_table(self):
        df = self.wb.sheets['Sheet4']['A1'].options(pd.DataFrame, expand='table').value
        df.index.name = None
        assert_frame_equal(df, data['df1'])
        self.assertIsNotNone(self.wb.sheets['Sheet4']['A1'].table)

    def test_var_operations(self):
        assert_array_equal(self.wb.sheets[1]['A1'].options(np.array, expand='table', ndim=2).value,
                           data['mydict']['df'][:1].values)

    def test_picture(self):
        self.assertEqual(self.wb.sheets[1].pictures[0].top, self.wb.sheets[1]['A17'].top)
        self.assertEqual(self.wb.sheets[1].pictures[0].left, self.wb.sheets[1]['A17'].left)

    def test_matplotlib(self):
        self.assertAlmostEqual(self.wb.sheets[1].pictures[1].top, self.wb.sheets[1]['B33'].top, places=2)
        self.assertAlmostEqual(self.wb.sheets[1].pictures[1].left, self.wb.sheets[1]['B33'].left, places=2)

    def test_used_range(self):
        self.assertEqual(self.wb.sheets[2]['B11'].value, 'This is text with a substringtest.')
        self.assertEqual(self.wb.sheets[2]['A1'].value, None)

    def test_different_vars_at_either_end(self):
        self.assertEqual(self.wb.sheets[0]['I1'].value, 'stringtest vs. stringtest')

    def test_shape_text(self):
        self.assertEqual(self.wb.sheets[4].shapes['TextBox 1'].text, 'This is no template. So the formatting should be left untouched.')
        self.assertEqual(self.wb.sheets[4].shapes['Oval 2'].text, 'This shows stringtest.')
        self.assertEqual(self.wb.sheets[4].shapes['TextBox 3'].text, 'This shows stringtest.')
        self.assertEqual(self.wb.sheets[4].shapes['TextBox 4'].text, 'stringtest')
        self.assertIsNone(self.wb.sheets[4].shapes['Oval 5'].text)

    def test_markdown_cell(self):
        self.assertEqual(self.wb.sheets['Sheet6']['A1'].value, 'Title\nText bold and italic\n\n• a first bullet\n• a second bullet\n\nAnother title\nthis has a line break\nnew line')

    def test_markdown_shape(self):
        self.assertEqual(self.wb.sheets['Sheet6'].shapes[0].text,
                         'Title\nText bold and italic\n\n• a first bullet\n• a second bullet\n\nAnother title\nthis has a line break\nnew line')


class TestBookSettings(unittest.TestCase):

    def tearDown(self):
        xw.Book('output.xlsx').app.quit()

    def test_update_links_false(self):
        wb = create_report('template_with_links.xlsx', 'output.xlsx', book_settings={'update_links': False}, **data)
        self.assertEqual(wb.sheets[0]['M1'].value, 'Text for update_links')

    @unittest.skipIf(sys.platform.startswith('darwin') is True, 'skip macOS')  # can't seem to make this test work on mac
    def test_update_links_true(self):
        wb = create_report('template_with_links.xlsx', 'output.xlsx', book_settings={'update_links': True}, **data)
        self.assertEqual(wb.sheets[0]['M1'].value, 'Updated Text for update_links')


class TestApp(unittest.TestCase):

    def test_app_instance(self):
        app = xw.App()
        wb = create_report('template_with_links.xlsx', 'output.xlsx', app=app, book_settings={'update_links': False}, **data)
        self.assertEqual(wb.sheets[0]['M1'].value, 'Text for update_links')
        wb.app.quit()


class TestFrames(unittest.TestCase):

    def tearDown(self):
        xw.Book('output.xlsx').app.quit()

    def test_one_frame(self):
        df = pd.DataFrame([[1., 2.], [3., 4.]],
                          columns=['c1', 'c2'],
                          index=['r1', 'r2'])
        wb = create_report('template_one_frame.xlsx', 'output.xlsx', df=df, title='MyTitle')
        for i in range(2):
            sheet = wb.sheets[i]
            self.assertEqual(sheet['A1'].value, 'MyTitle')
            self.assertEqual(sheet['A3'].value, 'PART ONE')
            self.assertEqual(sheet['A8'].value, 'PART TWO')
            if i == 0:
                assert_frame_equal(sheet['A4'].options(pd.DataFrame, expand='table').value, df)
                assert_frame_equal(sheet['A9'].options(pd.DataFrame, expand='table').value, df)
            elif i == 1:
                df_table1 = sheet['A4'].options(pd.DataFrame, expand='table').value
                df_table1.index.name = None
                df_table2 = sheet['A9'].options(pd.DataFrame, expand='table').value
                df_table2.index.name = None
                assert_frame_equal(df_table1, df)
                assert_frame_equal(df_table2, df)
            self.assertEqual(sheet['A3'].color, (0, 176, 240))
            self.assertEqual(sheet['A8'].color, (0, 176, 240))

    def test_two_frames(self):
        df1 = pd.DataFrame([[1., 2., 3.], [4., 5., 6.], [7., 8., 9.]],
                           columns=['c1', 'c2', 'c3'],
                           index=['r1', 'r2', 'r3'])
        df1.index.name = 'df1'

        df3 = pd.DataFrame([[1., 2., 3.], [4., 5., 6.], [7., 8., 9.], [10., 11., 12.], [13., 14., 15.]],
                           columns=['c1', 'c2', 'c3'],
                           index=['r1', 'r2', 'r3', 'r4', 'r5'])
        df3.index.name = 'df3'

        text = 'abcd'
        pic = Image(os.path.abspath('xlwings.jpg'))

        data = dict(df1=df1, df2='df2 dummy', df3=df3, text=text, pic=pic)
        wb = create_report('template_two_frames.xlsx', 'output.xlsx', **data)
        sheet = wb.sheets[0]
        # values
        assert_frame_equal(sheet['A1'].options(pd.DataFrame, expand='table').value, df3)
        self.assertEqual(sheet['A8'].value, 'df2 dummy')
        self.assertEqual(sheet['C10'].value, 'abcd')
        assert_frame_equal(sheet['A12'].options(pd.DataFrame, expand='table').value, df1)
        assert_frame_equal(sheet['A17'].options(pd.DataFrame, expand='table').value, df3)
        assert_frame_equal(sheet['A24'].options(pd.DataFrame, expand='table').value, df3)
        assert_frame_equal(sheet['A31'].options(pd.DataFrame, expand='table').value, df3)

        assert_frame_equal(sheet['F1'].options(pd.DataFrame, expand='table').value, df1)
        self.assertEqual(sheet['G6'].value, 'abcd')
        assert_frame_equal(sheet['F8'].options(pd.DataFrame, expand='table').value, df3)
        assert_frame_equal(sheet['F15'].options(pd.DataFrame, expand='table').value, df1)
        assert_frame_equal(sheet['F27'].options(pd.DataFrame, expand='table').value, df1)
        self.assertEqual(sheet['F32'].value, 'df2 dummy')
        assert_frame_equal(sheet['F34'].options(pd.DataFrame, expand='table').value, df3)

        # colors
        self.assertEqual(sheet['A2:D6'].color, (221, 235, 247))
        self.assertEqual(sheet['A13:D15'].color, (221, 235, 247))
        self.assertEqual(sheet['A18:D22'].color, (221, 235, 247))
        self.assertEqual(sheet['A25:D29'].color, (221, 235, 247))
        self.assertEqual(sheet['A32:D36'].color, (221, 235, 247))

        self.assertEqual(sheet['F2:I4'].color, (221, 235, 247))
        self.assertEqual(sheet['F9:I13'].color, (221, 235, 247))
        self.assertEqual(sheet['F16:I18'].color, (221, 235, 247))
        self.assertEqual(sheet['F28:I30'].color, (221, 235, 247))
        self.assertEqual(sheet['F35:I39'].color, (221, 235, 247))

        # borders
        # TODO: pending Border implementation in xlwings
        if sys.platform.startswith('darwin'):
            from appscript import k as kw
            for cell in ['A4', 'A14', 'D20', 'A28', 'D36', 'F4', 'H10', 'G17', 'G28', 'I36']:
                self.assertEqual(sheet[cell].api.get_border(which_border=kw.edge_top).properties().get(kw.line_style),
                                 kw.continuous)
                self.assertEqual(sheet[cell].api.get_border(which_border=kw.edge_bottom).properties().get(kw.line_style),
                                 kw.continuous)
        else:
            pass
            # TODO


class TestDataFrameFilters(unittest.TestCase):

    def tearDown(self):
        xw.Book('output.xlsx').app.quit()

    def test_df_filters(self):
        wb = create_report('template1.xlsx', 'output.xlsx', **data)
        self.assertEqual(wb.sheets['df_filters']['A1:E124'].value, wb.sheets['df_filters']['G1:K124'].value)

    def test_df_filters_in_frames(self):
        wb = create_report('df_filter_frame.xlsx', 'output.xlsx', **data)
        self.assertEqual(wb.sheets['Sheet1']['A1:E10'].value, wb.sheets['expected']['A1:E10'].value)
        self.assertEqual(wb.sheets['Sheet1']['A3'].color, wb.sheets['expected']['A3'].color)
        self.assertEqual(wb.sheets['Sheet1']['A4:A5'].color, wb.sheets['expected']['A4:A5'].color)
        self.assertIsNone(wb.sheets['Sheet1']['A6'].color)
        self.assertEqual(wb.sheets['Sheet1']['A7'].color, wb.sheets['expected']['A7'].color)
        self.assertEqual(wb.sheets['Sheet1']['A8'].color, wb.sheets['expected']['A8'].color)
        self.assertIsNone(wb.sheets['Sheet1']['A9'].color)

    def test_pic_filters(self):
        wb = create_report('template1.xlsx', 'output.xlsx', **data)
        self.assertEqual(wb.sheets['pic_filters'].pictures[0].width, 397)
        self.assertEqual(wb.sheets['pic_filters'].pictures[0].height, 139)
        self.assertEqual(wb.sheets['pic_filters'].pictures[1].width, 120)
        self.assertEqual(int(wb.sheets['pic_filters'].pictures[1].height), 42)
        self.assertEqual(int(wb.sheets['pic_filters'].pictures[2].width), 371)
        self.assertEqual(wb.sheets['pic_filters'].pictures[2].height, 130)
        self.assertEqual(int(wb.sheets['pic_filters'].pictures[3].width), 476)
        self.assertEqual(int(wb.sheets['pic_filters'].pictures[3].height), 166)


if __name__ == '__main__':
    unittest.main()
