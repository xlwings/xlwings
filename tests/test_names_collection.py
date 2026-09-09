"""Names filtering, scoped identity, and index shifts with stable and lazy engines."""

import sys
from types import SimpleNamespace

import pytest

from xlwings import base_classes
from xlwings._names import NameIndex
from xlwings.main import Names


class NativeName:
    def __init__(self, collection, key):
        self.collection = collection
        if isinstance(key, int):
            self.record = collection.api[key - 1]
        else:
            self.record = next(item for item in collection.api if item["name"] == key)

    @property
    def name(self):
        return self.record["name"]

    @property
    def refers_to(self):
        return self.record["formula"]

    @refers_to.setter
    def refers_to(self, value):
        self.record["formula"] = value

    @property
    def refers_to_range(self):
        # Reproduce #2569: Excel inserts a name while resolving a broken range.
        if not self.collection.inserted:
            self.collection.api.insert(
                0, {"name": "_xlfn.ANCHORARRAY", "formula": "=#NAME?"}
            )
            self.collection.inserted = True
        raise ValueError("Broken reference")

    def delete(self):
        self.collection.api.remove(self.record)


class NativeNames(base_classes.Names):
    def __init__(self, records):
        self.api = records
        self.inserted = False

    def __len__(self):
        return len(self.api)

    def __call__(self, key):
        if isinstance(key, str) and not self.contains(key):
            raise KeyError(key)
        return NativeName(self, key)

    def contains(self, key):
        return any(item["name"] == key for item in self.api)

    def add(self, name, refers_to):
        self.api.append({"name": name, "formula": refers_to})
        return self(name)


class LazyNativeName(NativeName):
    def __init__(self, collection, index):
        self.collection = collection
        self.index = index

    @property
    def record(self):
        index = self.index.resolve(
            self.collection.name_at_index,
            self.collection.name_strings,
            self.collection.sheet_names,
        )
        return self.collection.api[index - 1]


class LazyNativeNames(NativeNames):
    def __init__(self, records, sheets=None):
        super().__init__(records)
        self.sheets = (
            list(sheets)
            if sheets is not None
            else list(
                dict.fromkeys(
                    scope
                    for record in records
                    if (scope := NameIndex.split_name(record["name"])[0]) is not None
                )
            )
        )

    def sheet_names(self):
        return self.sheets

    def rename_sheet(self, old, new):
        self.sheets[self.sheets.index(old)] = new
        for record in self.api:
            scope, local = NameIndex.split_name(record["name"])
            if scope == old:
                quoted = "'" + new.replace("'", "''") + "'" if " " in new else new
                record["name"] = f"{quoted}!{local}"
        self.api.sort(
            key=lambda record: (not record["name"].startswith("_xlfn."), record["name"])
        )

    def name_at_index(self, index):
        if 1 <= index <= len(self.api):
            return self.api[index - 1]["name"]
        return None

    def name_strings(self):
        return [record["name"] for record in self.api]

    def __call__(self, key):
        if isinstance(key, str):
            key = self.name_strings().index(key) + 1
        return LazyNativeName(
            self, NameIndex(key, self.name_at_index(key), self.sheets)
        )

    def snapshot(self):
        return [
            (name, LazyNativeName(self, NameIndex(i, name, self.sheets)))
            for i, name in enumerate(self.name_strings(), 1)
        ]


@pytest.fixture(params=[NativeNames, LazyNativeNames], ids=["stable", "lazy"])
def native_names(request):
    return request.param


@pytest.mark.parametrize("prefix", ["", "'_xlfn.Scope!Name'!"])
def test_shared_names_filter_all_collection_operations(prefix, native_names):
    excluded = ["_xlfn.LAMBDA", "_xlpm.value", "_XLFN.UNIQUE", "_XLPM.Value"]
    included = [
        "UserRange",
        "HiddenConstant",
        "UserLambda",
        "BrokenUserName",
        "_xlfnCustom",
        "_xlpm_config",
    ]
    native = native_names(
        [
            {"name": prefix + name, "formula": "=#NAME?", "visible": False}
            for name in [excluded[0], *included[:2], *excluded[1:], *included[2:]]
        ]
    )
    names = Names(native)
    expected = [prefix + name for name in included]
    assert len(names) == names.count == len(included)
    assert [name.name for name in names] == expected
    assert [names[i].name for i in range(len(names))] == expected
    assert [names(i).name for i in range(1, len(names) + 1)] == expected
    assert names[-1].name == expected[-1]
    assert 0 in names and len(names) - 1 in names
    assert len(names) not in names
    assert names.contains(1) and names.contains(len(names))
    assert not names.contains(0) and not names.contains(len(names) + 1)
    with pytest.raises(IndexError):
        names[len(names)]
    with pytest.raises(IndexError):
        names(0)
    for name in expected:
        assert name in names and names.contains(name)
        assert names[name].name == name
    for name in excluded:
        key = prefix + name
        assert key not in names and not names.contains(key)
        with pytest.raises(KeyError):
            names[key]
    assert "_xlfn.LAMBDA" not in repr(names)
    # Low-level access remains available, including all native hidden names.
    assert names.api is native.api
    assert len(names.api) == len(included) + len(excluded)


def test_iteration_survives_internal_name_inserted_during_reference_resolution(
    native_names,
):
    native = native_names(
        [{"name": f"Range{i}", "formula": "=#REF!"} for i in range(1, 5)]
    )
    names = Names(native)
    seen = []
    for name in names:
        seen.append(name.name)
        if name.name == "Range2":
            with pytest.raises(ValueError, match="Broken reference"):
                name.refers_to_range
            # The handle must still refer to Range2 after the index shift.
            assert name.name == "Range2"
    assert seen == ["Range1", "Range2", "Range3", "Range4"]
    assert [name.name for name in names] == seen
    assert len(names) == 4
    assert native.api[0]["name"] == "_xlfn.ANCHORARRAY"


def test_filtered_indices_edit_and_delete_the_intended_name(native_names):
    native = native_names(
        [
            {"name": "_xlpm.value", "formula": "=#NAME?"},
            {"name": "First", "formula": "=1"},
            {"name": "_xlfn.LAMBDA", "formula": "=#NAME?"},
            {"name": "Second", "formula": "=2"},
        ]
    )
    names = Names(native)
    names[1].refers_to = "=20"
    assert names["Second"].refers_to == "=20"
    del names[0]
    assert [name.name for name in names] == ["Second"]
    names.add("Third", "=3")
    assert len(names) == 2
    assert names[1].name == "Third"
    # Never remove the native entries as a side effect of hiding them.
    assert native.contains("_xlfn.LAMBDA") and native.contains("_xlpm.value")


def test_duplicate_name_text_preserves_each_entry():
    native = NativeNames(
        [
            {"name": "_xlfn.LAMBDA", "formula": "=#NAME?"},
            {"name": "two", "formula": "=Sheet1!$A$1"},
            {"name": "two", "formula": "=Sheet2!$B$2"},
            {"name": "two", "formula": "=Sheet3!$C$3"},
        ]
    )
    names = Names(native)
    expected = [record["formula"] for record in native.api[1:]]
    assert [name.refers_to for name in names] == expected
    assert [names[i].refers_to for i in range(3)] == expected
    assert [names(i).refers_to for i in range(1, 4)] == expected
    assert names[-1].refers_to == expected[-1]
    names[1].refers_to = "=42"
    assert native.api[2]["formula"] == "=42"
    assert native.api[1]["formula"] == expected[0]
    del names[2]
    assert [name.refers_to for name in names] == [expected[0], "=42"]


def test_lazy_name_recovers_after_deletion_and_rejects_missing_entry():
    native = LazyNativeNames(
        [{"name": name, "formula": "=1"} for name in ["First", "Second", "Third"]]
    )
    first, second, third = list(Names(native))
    first.delete()
    assert third.name == "Third"  # The old index is now past the end.
    assert second.name == "Second"
    second.delete()
    with pytest.raises(KeyError, match="Second"):
        second.refers_to = "=42"
    assert third.refers_to == "=1"


def test_name_index_rescans_only_when_needed():
    from unittest.mock import Mock

    get_name = Mock(return_value="Range2")
    get_names = Mock(return_value=["_xlfn.ANCHORARRAY", "Range1", "Range2"])
    index = NameIndex(2, "Range2")
    assert index.resolve(get_name, get_names) == 2
    get_name.assert_called_once_with(2)
    get_names.assert_not_called()
    get_name.return_value = "Range1"
    assert index.resolve(get_name, get_names) == 3
    get_names.assert_called_once_with()
    get_name.return_value = "Range2"
    assert index.resolve(get_name, get_names) == 3
    assert get_names.call_count == 1


@pytest.mark.parametrize(
    "new_scope", ["Renamed", "'Renamed ! Sheet'", "'Owner''s Sheet'"]
)
def test_lazy_name_survives_sheet_rename(new_scope):
    native = LazyNativeNames(
        [
            {"name": "Sheet1!foo", "formula": "=1"},
            {"name": "Sheet2!foo", "formula": "=2"},
        ]
    )
    first, second = list(Names(native))
    native.rename_sheet("Sheet1", NameIndex.split_name(f"{new_scope}!foo")[0])
    assert first.name == f"{new_scope}!foo"
    first.refers_to = "=42"
    assert first.refers_to == "=42"
    assert second.refers_to == "=2"
    first.delete()
    assert second.name == "Sheet2!foo"


@pytest.mark.parametrize("local", ["Print_Area", "_FilterDatabase"])
@pytest.mark.parametrize("insert_internal", [False, True])
def test_scope_rename_tracks_worksheet_when_names_resort(local, insert_internal):
    native = LazyNativeNames(
        [
            {"name": f"{scope}!{local}", "formula": f"={i}"}
            for i, scope in enumerate(["Alpha", "Sheet1", "Zeta"], 1)
        ]
    )
    first, target, neighbor = list(Names(native))
    if insert_internal:
        native.api.insert(0, {"name": "_xlfn.ANCHORARRAY", "formula": "=#NAME?"})
    native.rename_sheet("Sheet1", "Zulu")
    assert target.name == f"Zulu!{local}"
    assert target.refers_to == "=2"
    target.refers_to = "=42"
    assert neighbor.name == f"Zeta!{local}"
    assert neighbor.refers_to == "=3"
    target.delete()
    assert first.refers_to == "=1"
    assert neighbor.refers_to == "=3"
    assert len(Names(native)) == 2


def test_independent_sheet_renames_preserve_each_scope():
    native = LazyNativeNames(
        [
            {"name": f"{scope}!foo", "formula": f"={i}"}
            for i, scope in enumerate(["Alpha", "Sheet1", "Zeta"], 1)
        ]
    )
    first, second, third = list(Names(native))
    native.rename_sheet("Alpha", "Gamma")
    native.rename_sheet("Sheet1", "Zulu")
    assert first.name == "Gamma!foo"
    assert second.name == "Zulu!foo"
    assert third.name == "Zeta!foo"
    assert [first.refers_to, second.refers_to, third.refers_to] == ["=1", "=2", "=3"]


def test_deleted_scoped_name_does_not_adopt_neighbor():
    native = LazyNativeNames(
        [
            {"name": f"{scope}!Print_Area", "formula": f"={i}"}
            for i, scope in enumerate(["Alpha", "Sheet1", "Zeta"], 1)
        ]
    )
    target = Names(native)[1]
    del native.api[1]
    with pytest.raises(KeyError, match="Sheet1!Print_Area"):
        target.refers_to = "=42"
    assert native.api[1]["formula"] == "=3"


def test_rename_with_worksheet_reorder_does_not_guess_scope():
    native = LazyNativeNames(
        [
            {"name": f"{scope}!Print_Area", "formula": f"={i}"}
            for i, scope in enumerate(["Alpha", "Sheet1", "Zeta"], 1)
        ]
    )
    target = Names(native)[1]
    native.rename_sheet("Sheet1", "Zulu")
    native.sheets = ["Alpha", "Zeta", "Zulu"]
    with pytest.raises(KeyError, match="Sheet1!Print_Area"):
        target.delete()
    assert len(native.api) == 3


def test_name_index_prefers_exact_match_over_changed_scope_at_old_index():
    index = NameIndex(1, "Sheet1!foo")
    names = ["Sheet2!foo", "Sheet1!foo"]
    assert index.resolve(lambda i: names[i - 1], lambda: names) == 2
    assert index.name == "Sheet1!foo"


@pytest.mark.parametrize(
    "old_name, current_name",
    [("Sheet1!foo", "foo"), ("foo", "Sheet1!foo"), ("Sheet1!foo", "Sheet2!bar")],
)
def test_name_index_rejects_changes_other_than_scope(old_name, current_name):
    index = NameIndex(1, old_name)
    with pytest.raises(KeyError):
        index.resolve(lambda i: current_name, lambda: [current_name])
    assert index.name == old_name


@pytest.mark.skipif(sys.platform != "darwin", reason="Requires the appscript engine")
@pytest.mark.parametrize("bulk_shadows", [False, True])
def test_mac_snapshot_preserves_shadowed_names(bulk_shadows):
    from xlwings._xlmac import Names as MacNames, kw

    class NativeCollection:
        def __init__(self):
            self.records = [
                {"name": "foo", "formula": "=1"},
                {"name": "Sheet1!foo", "formula": "=2"},
            ]
            self.bulk_reads = 0
            self.index_reads = 0
            self.sheet_reads = 0
            self.name = SimpleNamespace(get=self.get_names)

        def get_sheet_names(self):
            self.sheet_reads += 1
            return ["Sheet1"]

        def get_names(self):
            self.bulk_reads += 1
            names = [record["name"] for record in self.records]
            if bulk_shadows and "foo" in names and "Sheet1!foo" in names:
                names[names.index("foo")] = "Sheet1!foo"
            return names

        def __getitem__(self, key):
            assert isinstance(
                key, int
            ), "By-name lookup would select the shadowing name"
            return NativeReference(self, key)

    class NativeReference:
        def __init__(self, collection, index):
            self.collection = collection
            self.index = index
            self.name = SimpleNamespace(get=self.get_name, set=self.set_name)
            self.references = SimpleNamespace(
                set=lambda value: self.record.update(formula=value)
            )

        @property
        def record(self):
            return self.collection.records[self.index - 1]

        def get_name(self):
            self.collection.index_reads += 1
            return self.record["name"]

        def set_name(self, value):
            if value == "bad name" or value in self.collection.get_names():
                # Excel may reject a rename without raising an Apple Event error.
                return
            scope = self.record["name"].rpartition("!")[0]
            self.record["name"] = f"{scope}!{value}" if scope else value
            self.collection.records.sort(key=lambda record: record["name"])

        def properties(self):
            return {kw.references: self.record["formula"]}

        def delete(self):
            self.collection.records.remove(self.record)

    native = NativeCollection()
    # The live test covers Excel's mutation scope. This fake tests index tracking.
    active = SimpleNamespace(
        names=SimpleNamespace(_name_strings=lambda: []),
        xl=SimpleNamespace(exists=lambda: True),
    )
    parent = SimpleNamespace(
        book=SimpleNamespace(
            sheets=SimpleNamespace(active=active),
            xl=SimpleNamespace(
                worksheets=SimpleNamespace(
                    name=SimpleNamespace(get=native.get_sheet_names)
                )
            ),
        )
    )
    names = Names(MacNames(parent=parent, xl=native))
    iterator = iter(names)
    assert native.bulk_reads == 1
    assert native.sheet_reads == 1
    assert native.index_reads == (2 if bulk_shadows else 0)
    book_name, sheet_name = iterator
    assert book_name.refers_to == "=1"
    assert sheet_name.refers_to == "=2"
    assert native.bulk_reads == 1
    native.records.insert(0, {"name": "_xlfn.ANCHORARRAY", "formula": "=#NAME?"})
    book_name.refers_to = "=42"
    assert native.records[1]["formula"] == "=42"
    assert sheet_name.refers_to == "=2"
    assert [name.name for name in names] == ["foo", "Sheet1!foo"]
    book_name.name = "zzz"
    assert book_name.name == "zzz"
    sheet_name.name = "renamed"
    assert sheet_name.name == "Sheet1!renamed"
    assert sheet_name.refers_to == "=2"
    from xlwings import XlwingsError

    for rejected in ["bad name", "Sheet1!renamed"]:
        with pytest.raises(XlwingsError, match="did not rename"):
            book_name.name = rejected
        assert book_name.name == "zzz"
        assert book_name.refers_to == "=42"
        assert sheet_name.refers_to == "=2"
    book_name.delete()
    assert [record["name"] for record in native.records] == [
        "Sheet1!renamed",
        "_xlfn.ANCHORARRAY",
    ]
