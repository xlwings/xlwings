import os
import sys
import unittest

import xlwings as xw

from .common import TestBase, this_dir

try:
    import numpy as np
except ImportError:
    np = None

try:
    import matplotlib as mpl
    import matplotlib.pyplot as plt
except ImportError:
    mpl = None

try:
    import PIL
except ImportError:
    PIL = None

if sys.version_info[0] >= 3 and sys.version_info[1] >= 6:
    import pathlib
else:
    pathlib = None

try:
    import plotly.graph_objects as plotly_go
except ImportError:
    plotly_go = None


class TestShape(TestBase):
    def test_name(self):
        filename = os.path.join(this_dir, "sample_picture.png")
        self.wb1.sheets[0].pictures.add(filename, name="pic1")

        sh = self.wb1.sheets[0].shapes[0]
        self.assertEqual(sh.name, "pic1")
        sh.name = "yoyoyo"
        self.assertEqual(sh.name, "yoyoyo")

    @unittest.skipIf(pathlib is None, "pathlib unavailable")
    def test_name_pathlib(self):
        filename = pathlib.Path(this_dir) / "sample_picture.png"
        self.wb1.sheets[0].pictures.add(filename, name="pic1")

        sh = self.wb1.sheets[0].shapes[0]
        self.assertEqual(sh.name, "pic1")
        sh.name = "yoyoyo"
        self.assertEqual(sh.name, "yoyoyo")

    def test_coordinates(self):
        filename = os.path.join(this_dir, "sample_picture.png")
        self.wb1.sheets[0].pictures.add(
            filename, name="pic1", left=0, top=0, width=200, height=100
        )

        sh = self.wb1.sheets[0].shapes[0]
        for a, init, neu in (
            ("left", 0, 50),
            ("top", 0, 50),
            ("width", 200, 150),
            ("height", 100, 160),
        ):
            self.assertEqual(getattr(sh, a), init)
            setattr(sh, a, neu)
            self.assertEqual(getattr(sh, a), neu)

    def test_picture_object(self):
        filename = os.path.join(this_dir, "sample_picture.png")
        self.wb1.sheets[0].pictures.add(filename, name="pic1")

        self.assertEqual(
            self.wb1.sheets[0].shapes[0], self.wb1.sheets[0].shapes["pic1"]
        )

    def test_delete(self):
        filename = os.path.join(this_dir, "sample_picture.png")
        self.wb1.sheets[0].pictures.add(filename, name="pic1")
        self.assertTrue("pic1" in self.wb1.sheets[0].shapes)
        self.wb1.sheets[0].shapes[0].delete()
        self.assertFalse("pic1" in self.wb1.sheets[0].shapes)

    def test_type(self):
        filename = os.path.join(this_dir, "sample_picture.png")
        self.wb1.sheets[0].pictures.add(filename, name="pic1")
        self.assertEqual(self.wb1.sheets[0].shapes[0].type, "picture")

    def test_scale_width(self):
        filename = os.path.join(this_dir, "sample_picture.png")
        pic = self.wb1.sheets[0].pictures.add(filename, name="pic1")
        w, h = int(pic.width), int(pic.height)
        self.wb1.sheets[0].shapes["pic1"].scale_width(factor=2)
        self.assertEqual(int(pic.width), w * 2)
        self.assertEqual(int(pic.height), h * 2)

    def test_scale_height(self):
        filename = os.path.join(this_dir, "sample_picture.png")
        pic = self.wb1.sheets[0].pictures.add(filename, name="pic1")
        w, h = int(pic.width), int(pic.height)
        self.wb1.sheets[0].shapes["pic1"].scale_height(factor=2)
        self.assertEqual(int(pic.width), w * 2)
        self.assertEqual(int(pic.height), h * 2)


class TestPicture(TestBase):
    def test_two_books(self):
        filename = os.path.join(this_dir, "sample_picture.png")
        pic1 = self.wb1.sheets[0].pictures.add(filename, name="pic1")
        pic2 = self.wb2.sheets[0].pictures.add(filename, name="pic1")
        self.assertEqual(pic1.name, "pic1")
        self.assertEqual(pic2.name, "pic1")

    def test_name(self):
        filename = os.path.join(this_dir, "sample_picture.png")
        pic = self.wb1.sheets[0].pictures.add(filename, name="pic1")
        self.assertEqual(pic.name, "pic1")

        pic.name = "pic_new"
        self.assertEqual(pic.name, "pic_new")

    def test_left(self):
        filename = os.path.join(this_dir, "sample_picture.png")
        pic = self.wb1.sheets[0].pictures.add(filename, name="pic1")
        self.assertEqual(pic.left, 0)

        pic.left = 20
        self.assertEqual(pic.left, 20)

    def test_top(self):
        filename = os.path.join(this_dir, "sample_picture.png")
        pic = self.wb1.sheets[0].pictures.add(filename, name="pic1")
        self.assertEqual(pic.left, 0)

        pic.top = 20
        self.assertEqual(pic.top, 20)

    def test_width(self):
        filename = os.path.join(this_dir, "sample_picture.png")
        pic = self.wb1.sheets[0].pictures.add(filename, name="pic1")
        self.assertEqual(int(pic.width), 30)
        pic.width = 50
        self.assertEqual(pic.width, 50)

    def test_picture_object(self):
        filename = os.path.join(this_dir, "sample_picture.png")
        pic = self.wb1.sheets[0].pictures.add(filename, name="pic1")
        self.assertEqual(pic.name, self.wb1.sheets[0].pictures["pic1"].name)

    def test_height(self):
        filename = os.path.join(this_dir, "sample_picture.png")
        pic = self.wb1.sheets[0].pictures.add(filename, name="pic1")
        self.assertEqual(int(pic.height), 30)
        pic.height = 50
        self.assertEqual(int(pic.height), 50)

    def test_delete(self):
        filename = os.path.join(this_dir, "sample_picture.png")
        pic = self.wb1.sheets[0].pictures.add(filename, name="pic1")
        self.assertTrue("pic1" in self.wb1.sheets[0].pictures)
        pic.delete()
        self.assertFalse("pic1" in self.wb1.sheets[0].pictures)

    def test_duplicate(self):
        with self.assertRaises(xw.ShapeAlreadyExists):
            filename = os.path.join(this_dir, "sample_picture.png")
            self.wb1.sheets[0].pictures.add(filename, name="pic1")
            self.wb1.sheets[0].pictures.add(filename, name="pic1")

    def test_picture_update(self):
        filename = os.path.join(this_dir, "sample_picture.png")
        pic1 = self.wb1.sheets[0].pictures.add(filename, name="pic1")
        pic1.update(filename)

    @unittest.skipIf(pathlib is None, "pathlib unavailable")
    def test_picture_update_pathlib(self):
        filename = pathlib.Path(this_dir) / "sample_picture.png"
        pic1 = self.wb1.sheets[0].pictures.add(filename, name="pic1")
        pic1.update(filename)

    def test_picture_auto_update(self):
        filename = os.path.join(this_dir, "sample_picture.png")
        self.wb1.sheets[0].pictures.add(filename, name="pic1", update=True)
        self.wb1.sheets[0].pictures.add(filename, name="pic1", update=True)
        self.assertEqual(len(self.wb1.sheets[0].pictures), 1)

    def test_picture_auto_update_without_name(self):
        with self.assertRaises(ValueError):
            filename = os.path.join(this_dir, "sample_picture.png")
            self.wb1.sheets[0].pictures.add(filename, update=True)

    def test_picture_index(self):
        filename = os.path.join(this_dir, "sample_picture.png")
        self.wb1.sheets[0].pictures.add(filename, name="pic1")
        self.assertEqual(
            self.wb1.sheets[0].pictures[0], self.wb1.sheets[0].pictures["pic1"]
        )
        self.assertEqual(self.wb1.sheets[0].pictures(1), self.wb1.sheets[0].pictures[0])

    def test_len(self):
        filename = os.path.join(this_dir, "sample_picture.png")
        self.wb1.sheets[0].pictures.add(filename, name="pic1")
        self.wb1.sheets[0].pictures.add(filename, name="pic2")
        self.assertEqual(len(self.wb1.sheets[0].pictures), 2)

    def test_iter(self):
        filename = os.path.join(this_dir, "sample_picture.png")
        names = ["pic1", "pic2"]
        self.wb1.sheets[0].pictures.add(filename, name=names[0])
        self.wb1.sheets[0].pictures.add(filename, name=names[1])
        for ix, pic in enumerate(self.wb1.sheets[0].pictures):
            self.assertEqual(self.wb1.sheets[0].pictures[ix].name, names[ix])

    def test_contains(self):
        filename = os.path.join(this_dir, "sample_picture.png")
        self.wb1.sheets[0].pictures.add(filename, name="pic 1")
        self.assertTrue("pic 1" in self.wb1.sheets[0].pictures)


@unittest.skipIf(mpl is None, "matplotlib missing")
class TestMatplotlib(TestBase):
    def test_add_no_name(self):
        fig = plt.figure()
        plt.plot([-1, 1, -2, 2, -3, 3, 2])
        self.wb1.sheets[0].pictures.add(fig)
        self.assertEqual(len(self.wb1.sheets[0].pictures), 1)

    def test_add_with_name(self):
        fig = plt.figure()
        plt.plot([-1, 1, -2, 2, -3, 3, 2])
        self.wb1.sheets[0].pictures.add(fig, name="Test1")
        self.assertEqual(self.wb1.sheets[0].pictures[0].name, "Test1")


@unittest.skipIf(plotly_go is None, "plotly missing")
class TestPlotly(TestBase):
    def get_plotly_fig(self):
        N = 100
        x = np.random.rand(N)
        y = np.random.rand(N)
        colors = np.random.rand(N)
        sz = np.random.rand(N) * 30

        fig = plotly_go.Figure()
        fig.add_trace(
            plotly_go.Scatter(
                x=x,
                y=y,
                mode="markers",
                marker=plotly_go.scatter.Marker(
                    size=sz, color=colors, opacity=0.6, colorscale="Viridis"
                ),
            )
        )
        return fig

    def test_add_no_name(self):
        self.wb1.sheets[0].pictures.add(self.get_plotly_fig())
        self.assertEqual(len(self.wb1.sheets[0].pictures), 1)

    def test_add_with_name(self):
        self.wb1.sheets[0].pictures.add(self.get_plotly_fig(), name="Test1")
        self.assertEqual(self.wb1.sheets[0].pictures[0].name, "Test1")


class TestCharts(TestBase):
    def test_add_properties(self):
        sht = self.wb1.sheets[0]
        sht.range("A1").value = [["one", "two"], [1.1, 2.2]]

        self.assertEqual(len(sht.charts), 0)
        chart = sht.charts.add()
        self.assertEqual(len(sht.charts), 1)

        chart.name = "My Chart"
        chart.set_source_data(sht.range("A1").expand("table"))
        chart.chart_type = "line"

        self.assertEqual("My Chart", chart.name)
        self.assertEqual(sht.charts[0].chart_type, "line")

        chart.chart_type = "pie"
        self.assertEqual(sht.charts[0].chart_type, "pie")

        for a in ("left", "top", "width", "height"):
            setattr(chart, a, 400)
            self.assertEqual(getattr(sht.charts[0], a), 400)
            setattr(sht.charts[0], a, 500)
            self.assertEqual(getattr(chart, a), 500)

        chart.delete()
        self.assertEqual(sht.charts.count, 0)


class TestChartFormatting(TestBase):
    def test_collection_parent(self):
        sht = self.wb1.sheets[0]
        try:
            self.assertEqual(sht.charts.parent, sht)
        except NotImplementedError:
            self.fail("Charts.parent must be implemented on desktop engines")

    def _chart(self):
        sht = self.wb1.sheets[0]
        sht.range("A1").value = [["x", "a", "b"], [1, 10, 20], [2, 30, 40]]
        return sht, sht.charts.add(source=sht.range("A1:C3"), chart_type="line")

    def test_title(self):
        sht, chart = self._chart()
        self.assertIsNone(chart.title)
        chart.title = "Sales"
        self.assertEqual(sht.charts[0].title, "Sales")
        chart.title = None
        self.assertIsNone(sht.charts[0].title)
        with self.assertRaises(ValueError):
            chart.title = 1

    def test_legend(self):
        sht, chart = self._chart()
        chart.legend.visible = False
        self.assertFalse(sht.charts[0].legend.visible)
        self.assertIsNone(sht.charts[0].legend.position)
        for position in ["top", "bottom", "left", "right", "corner"]:
            chart.legend.position = position
            self.assertTrue(sht.charts[0].legend.visible)
            self.assertEqual(sht.charts[0].legend.position, position)
        # position then hide
        chart.legend.position = "bottom"
        chart.legend.visible = False
        self.assertFalse(chart.legend.visible)
        self.assertIsNone(chart.legend.position)
        # hide then position
        chart.legend.visible = False
        chart.legend.position = "top"
        self.assertTrue(chart.legend.visible)
        self.assertEqual(chart.legend.position, "top")
        with self.assertRaises(ValueError):
            chart.legend.position = "middle"

    def test_legend_retained_after_rename(self):
        sht, chart = self._chart()
        legend = chart.legend
        chart.name = "Renamed"
        legend.position = "bottom"
        self.assertEqual(sht.charts["Renamed"].legend.position, "bottom")
        legend.visible = False
        self.assertFalse(sht.charts["Renamed"].legend.visible)
        self.assertFalse(chart.legend.visible)
        chart.title = "After"
        self.assertEqual(sht.charts["Renamed"].title, "After")

    def test_plot_by(self):
        sht, chart = self._chart()
        chart.set_source_data(sht.range("A1:C3"), plot_by="rows")
        self.assertEqual(sht.charts[0].plot_by, "rows")
        chart.plot_by = "columns"
        self.assertEqual(sht.charts[0].plot_by, "columns")
        with self.assertRaises(ValueError):
            chart.plot_by = "cols"

    def test_style(self):
        sht, chart = self._chart()
        chart.style = 10
        self.assertEqual(sht.charts[0].style, 10)
        for value in [0, 49, 12.0, True, "12", None]:
            with self.assertRaises(ValueError):
                chart.style = value

    def test_add_with_options(self):
        sht = self.wb1.sheets[0]
        sht.range("A1").value = [["x", "a", "b"], [1, 10, 20], [2, 30, 40]]
        chart = sht.charts.add(
            source=sht.range("A1:C3"),
            chart_type="pie",
            plot_by="rows",
            name="MyChart",
            width=300,
            height=200,
        )
        self.assertEqual(chart.name, "MyChart")
        self.assertEqual(sht.charts["MyChart"].chart_type, "pie")
        self.assertEqual(sht.charts["MyChart"].plot_by, "rows")
        self.assertEqual(chart.width, 300)
        self.assertEqual(chart.height, 200)
        try:
            with self.assertRaises(xw.ShapeAlreadyExists):
                sht.charts.add(name="MyChart")
        except NotImplementedError:
            # TestBase otherwise turns a missing implementation into a skip.
            self.fail("Duplicate chart names must raise ShapeAlreadyExists")

    def test_add_anchor(self):
        sht = self.wb1.sheets[0]
        chart = sht.charts.add(anchor=sht.range("D5"))
        self.assertEqual(chart.top, sht.range("D5").top)
        self.assertEqual(chart.left, sht.range("D5").left)
        with self.assertRaises(ValueError):
            sht.charts.add(left=10, anchor=sht.range("D5"))
        with self.assertRaises(ValueError):
            sht.charts.add(top=10, anchor=sht.range("D5"))

    def test_add_plot_by_requires_source(self):
        with self.assertRaises(ValueError):
            self.wb1.sheets[0].charts.add(plot_by="rows")
        with self.assertRaises(ValueError):
            self.wb1.sheets[0].charts.add(chart_type="nonsense")


class TestChartSheet(TestBase):
    """Chart sheets aren't reachable via the public collections (Chart.parent
    assumes a worksheet), so they're created natively and wrapped directly."""

    def _source(self):
        sht = self.wb1.sheets[0]
        sht.range("A1").value = [["x", "a", "b"], [1, 10, 20], [2, 30, 40]]
        return sht

    @unittest.skipUnless(sys.platform.startswith("darwin"), "macOS only")
    def test_chart_sheet_mac(self):
        from appscript import k as kw

        from xlwings._xlmac import Chart as MacChart

        sht = self._source()
        embedded = sht.charts.add(source=sht.range("A1:C3"), chart_type="line")
        embedded.api[1].chart_location(where=kw.location_as_new_sheet, name="CS")
        chart = xw.Chart(impl=MacChart(self.wb1.impl, "CS"))
        self.assertEqual(chart.name, "CS")
        self.assertEqual(chart.chart_type, "line")
        chart.title = "Title"
        chart.name = "Renamed"
        self.assertEqual(chart.name, "Renamed")
        self.assertEqual(chart.title, "Title")
        chart.style = 3
        self.assertEqual(chart.style, 3)
        chart.legend.position = "bottom"
        self.assertEqual(chart.legend.position, "bottom")
        with self.assertRaises(Exception):
            chart.left
        n_chart_sheets = self.wb1.api.count(each=kw.chart_sheet)
        chart.delete()
        self.assertEqual(self.wb1.api.count(each=kw.chart_sheet), n_chart_sheets - 1)

    @unittest.skipUnless(sys.platform.startswith("win"), "Windows only")
    def test_chart_sheet_win(self):
        from xlwings._xlwindows import Chart as WinChart

        sht = self._source()
        embedded = sht.charts.add(source=sht.range("A1:C3"), chart_type="line")
        embedded.api[1].Location(1, "CS")  # xlLocationAsNewSheet
        chart = xw.Chart(impl=WinChart(xl=self.wb1.api.Charts("CS")))
        self.assertEqual(chart.name, "CS")
        self.assertEqual(chart.chart_type, "line")
        chart.title = "Title"
        self.assertEqual(chart.title, "Title")
        chart.name = "Renamed"
        self.assertEqual(chart.name, "Renamed")
        chart.legend.position = "bottom"
        self.assertEqual(chart.legend.position, "bottom")
        n_chart_sheets = self.wb1.api.Charts.Count
        self.wb1.app.display_alerts = False
        try:
            chart.delete()
        finally:
            self.wb1.app.display_alerts = True
        self.assertEqual(self.wb1.api.Charts.Count, n_chart_sheets - 1)


class TestChart(TestBase):
    def test_len(self):
        self.wb1.sheets[0].charts.add()
        self.assertEqual(len(self.wb1.sheets[0].charts), 1)

    def test_count(self):
        self.wb1.sheets[0].charts.add()
        self.assertEqual(
            len(self.wb1.sheets[0].charts), self.wb1.sheets[0].charts.count
        )


if __name__ == "__main__":
    unittest.main()
