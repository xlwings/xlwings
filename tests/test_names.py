import sys
import unittest

from .common import TestBase


class TestNames(TestBase):
    def test_get_names_index(self):
        self.wb1.sheets[0].range("B2:D10").name = "test1"
        self.wb1.sheets[0].range("A1").name = "test2"
        self.assertEqual(self.wb1.names(1).name, "test1")
        self.assertEqual(self.wb1.names[1].name, "test2")

    @unittest.skipUnless(sys.platform == "darwin", "macOS scoped-name regression")
    def test_mac_shadowed_workbook_name_by_index(self):
        sheet = self.wb1.sheets[0]
        self.wb1.names.add("foo", "=Sheet1!$A$1")
        self.wb1.names.add("Sheet1!foo", "=Sheet1!$B$1")
        expected = {"foo": sheet["A1"], "Sheet1!foo": sheet["B1"]}
        expected_sheets = [s.name for s in self.wb1.sheets]
        sheet.activate()
        names = self.wb1.names
        self.assertEqual({name.name for name in names}, set(expected))
        self.assertEqual(len(names), 2)
        for i in range(len(names)):
            self.assertEqual(names[i].refers_to_range, expected[names[i].name])
            self.assertEqual(names(i + 1).refers_to_range, expected[names[i].name])
        for name in names:
            self.assertEqual(name.refers_to_range, expected[name.name])
        retained = list(names)
        names.add("aaa", "=1")
        for name in retained:
            self.assertEqual(name.refers_to_range, expected[name.name])
        names["aaa"].delete()
        book_index = next(i for i, name in enumerate(names) if name.name == "foo")
        sheet["D5"].select()
        control = names.add("relative_control", "=A2")
        expected_relative = control.refers_to_range
        control.delete()
        names[book_index].refers_to = "=A2"
        self.assertEqual(names[book_index].refers_to_range, expected_relative)
        self.assertEqual(sheet.names[0].refers_to_range, sheet["B1"])
        handle = names[book_index]
        handle.name = "renamed"
        self.assertEqual(handle.name, "renamed")
        self.assertEqual(handle.refers_to_range, expected_relative)
        handle.name = "foo"
        self.assertEqual(handle.name, "foo")
        self.assertEqual(handle.refers_to_range, expected_relative)
        book_index = next(i for i, name in enumerate(names) if name.name == "foo")
        del names[book_index]
        self.assertEqual(len(names), 1)
        self.assertEqual(names[0].name, "Sheet1!foo")
        self.assertEqual(names[0].refers_to_range, sheet["B1"])
        self.assertEqual([s.name for s in self.wb1.sheets], expected_sheets)
        self.assertEqual(self.wb1.sheets.active.name, sheet.name)
        self.assertEqual(self.wb1.app.selection, sheet["D5"])

    @unittest.skipUnless(sys.platform == "darwin", "macOS scoped-name regression")
    def test_mac_held_name_survives_sheet_rename(self):
        sheet = self.wb1.sheets[0]
        self.wb1.names.add("foo", "=Sheet1!$A$1")
        self.wb1.names.add("Sheet1!foo", "=Sheet1!$B$1")
        handles = {name.name: name for name in self.wb1.names}
        sheet.name = "Renamed"
        local = handles["Sheet1!foo"]
        self.assertEqual(local.name, "Renamed!foo")
        self.assertEqual(local.refers_to_range, sheet["B1"])
        self.assertEqual(handles["foo"].refers_to_range, sheet["A1"])
        local.refers_to = "=Renamed!$B$2"
        self.assertEqual(local.refers_to_range, sheet["B2"])
        local.delete()
        self.assertEqual(handles["foo"].refers_to_range, sheet["A1"])

    @unittest.skipUnless(sys.platform == "darwin", "macOS scoped-name regression")
    def test_mac_held_name_survives_sorted_scope_rename(self):
        alpha = self.wb1.sheets[0]
        alpha.name = "Alpha"
        middle = self.wb1.sheets.add(name="Sheet1", after=alpha)
        zeta = self.wb1.sheets.add(name="Zeta", after=middle)
        for sheet in [alpha, middle, zeta]:
            self.wb1.names.add(f"{sheet.name}!Print_Area", f"={sheet.name}!$A$1")
        names = self.wb1.names
        handles = {name.name: name for name in names}
        target = handles["Sheet1!Print_Area"]
        middle.name = "Zulu"
        self.assertEqual(target.name, "Zulu!Print_Area")
        self.assertEqual(target.refers_to_range, middle["A1"])
        target.refers_to = "=Zulu!$B$2"
        self.assertEqual(target.refers_to_range, middle["B2"])
        self.assertEqual(handles["Zeta!Print_Area"].refers_to_range, zeta["A1"])
        target.delete()
        self.assertEqual(handles["Alpha!Print_Area"].refers_to_range, alpha["A1"])
        self.assertEqual(handles["Zeta!Print_Area"].refers_to_range, zeta["A1"])
        self.assertEqual(len(names), 2)

    @unittest.skipUnless(sys.platform == "darwin", "macOS chart-sheet regression")
    def test_mac_mutations_with_chart_sheet_active(self):
        from xlwings._xlmac import kw

        self.wb1.names.add("foo", "=Sheet1!$A$1")
        name = self.wb1.names[0]
        chart = self.wb1.api.make(new=kw.chart_sheet, at=self.wb1.api.sheets[1].before)
        chart_name = chart.name.get()
        chart.activate_object()
        worksheet_names = [sheet.name for sheet in self.wb1.sheets]
        name.refers_to = "=Sheet1!$B$1"
        self.assertEqual(name.refers_to_range, self.wb1.sheets["Sheet1"]["B1"])
        self.assertEqual(self.wb1.api.active_sheet.name.get(), chart_name)
        name.name = "renamed"
        self.assertEqual(name.name, "renamed")
        self.assertEqual(self.wb1.api.active_sheet.name.get(), chart_name)
        name.delete()
        self.assertEqual(len(self.wb1.names), 0)
        self.assertEqual(self.wb1.api.active_sheet.name.get(), chart_name)
        self.assertEqual([sheet.name for sheet in self.wb1.sheets], worksheet_names)

    def test_names_contain(self):
        self.wb1.sheets[0].range("B2:D10").name = "test1"
        self.assertTrue("test1" in self.wb1.names)

    def test_len(self):
        self.wb1.sheets[0].range("B2:D10").name = "test1"
        self.wb1.sheets[0].range("A1").name = "test2"
        self.assertEqual(len(self.wb1.names), 2)

    def test_count(self):
        self.assertEqual(len(self.wb1.names), self.wb1.names.count)

    def test_names_iter(self):
        self.wb1.sheets[0].range("B2:D10").name = "test1"
        self.wb1.sheets[0].range("A1").name = "test2"
        for ix, n in enumerate(self.wb1.names):
            if ix == 0:
                self.assertEqual(n.name, "test1")
            if ix == 1:
                self.assertEqual(n.name, "test2")

    def test_get_inexisting_name(self):
        self.assertIsNone(self.wb1.sheets[0].range("A1").name)

    def test_get_set_named_range(self):
        self.wb1.sheets[0].range("A100").name = "test1"
        self.assertEqual(self.wb1.sheets[0].range("A100").name.name, "test1")

        self.wb1.sheets[0].range("A200:B204").name = "test2"
        self.assertEqual(self.wb1.sheets[0].range("A200:B204").name.name, "test2")

    def test_delete_named_item1(self):
        self.wb1.sheets[0].range("B10:C11").name = "to_be_deleted"
        self.assertEqual(
            self.wb1.sheets[0].range("to_be_deleted").name.name, "to_be_deleted"
        )

        del self.wb1.names["to_be_deleted"]
        self.assertIsNone(self.wb1.sheets[0].range("B10:C11").name)

    def test_delete_named_item2(self):
        self.wb1.sheets[0].range("B10:C11").name = "to_be_deleted"
        self.assertEqual(
            self.wb1.sheets[0].range("to_be_deleted").name.name, "to_be_deleted"
        )

        self.wb1.names["to_be_deleted"].delete()
        self.assertIsNone(self.wb1.sheets[0].range("B10:C11").name)

    def test_delete_named_item3(self):
        self.wb1.sheets[0].range("B10:C11").name = "to_be_deleted"
        self.assertEqual(
            self.wb1.sheets[0].range("to_be_deleted").name.name, "to_be_deleted"
        )

        self.wb1.sheets[0].range("to_be_deleted").name.delete()
        self.assertIsNone(self.wb1.sheets[0].range("B10:C11").name)

    def test_names_collection(self):
        self.wb1.sheets[0].range("A1").name = "name1"
        self.wb1.sheets[0].range("A2").name = "name2"
        self.assertTrue("name1" in self.wb1.names and "name2" in self.wb1.names)

        self.wb1.sheets[0].range("A3").name = "name3"
        self.assertTrue(
            "name1" in self.wb1.names
            and "name2" in self.wb1.names
            and "name3" in self.wb1.names
        )

    def test_sheet_scope(self):
        self.wb2.sheets[0].range("B2:C3").name = "Sheet1!sheet_scope1"
        self.wb2.sheets[0].range("sheet_scope1").value = [[1.0, 2.0], [3.0, 4.0]]
        self.assertEqual(
            self.wb2.sheets[0].range("B2:C3").value, [[1.0, 2.0], [3.0, 4.0]]
        )
        with self.assertRaises(Exception):
            self.wb2.sheets[1].range("sheet_scope1").value

    def test_workbook_scope(self):
        self.wb1.sheets[0].range("A1").name = "test1"
        self.wb1.sheets[0].range("test1").value = 123.0
        self.assertEqual(self.wb1.names["test1"].refers_to_range.value, 123.0)

    def test_contains_name(self):
        self.wb1.sheets[0].range("A1").name = "test1"
        self.assertTrue(self.wb1.names.contains("test1"))
        self.assertFalse(self.wb1.names.contains("test2"))

    def test_wb_names_add(self):
        self.wb1.names.add("test1", "=Sheet1!$A$1:$B$3")
        self.assertEqual(self.wb1.sheets[0].range("A1:B3").name.name, "test1")

    def test_sht_names_add(self):
        self.wb1.sheets[0].names.add("test1", "=Sheet1!$A$1:$B$3")
        self.assertEqual(self.wb1.sheets[0].range("A1:B3").name.name, "Sheet1!test1")

    def test_refers_to_range(self):
        self.wb1.sheets[0].range("B2:D10").name = "test1"
        self.assertEqual(
            self.wb1.sheets[0].range("B2:D10"),
            self.wb1.sheets[0].range("B2:D10").name.refers_to_range,
        )

    def test_refers_to_range_sheet_with_spaces(self):
        # This will cause quotes around sheet reference which caused a bug on mac
        self.wb1.sheets[0].name = "She et1"
        self.wb1.sheets[0].range("B2:D10").name = "test1"
        self.assertEqual(
            self.wb1.sheets["She et1"].range("B2:D10"),
            self.wb1.names["test1"].refers_to_range,
        )


if __name__ == "__main__":
    unittest.main()
