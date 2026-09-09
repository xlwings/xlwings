"""Shared Names contracts, including native collections with lazy index handles."""

import pytest

from xlwings.main import Names


class NativeName:
    def __init__(self, collection, key):
        self.collection = collection
        self.key = key

    @property
    def record(self):
        if isinstance(self.key, int):
            return self.collection.api[self.key - 1]
        return next(item for item in self.collection.api if item["name"] == self.key)

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


class NativeNames:
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


@pytest.mark.parametrize("prefix", ["", "'_xlfn.Scope!Name'!"])
def test_shared_names_filter_all_collection_operations(prefix):
    excluded = ["_xlfn.LAMBDA", "_xlpm.value", "_XLFN.UNIQUE", "_XLPM.Value"]
    included = [
        "UserRange",
        "HiddenConstant",
        "UserLambda",
        "BrokenUserName",
        "_xlfnCustom",
        "_xlpm_config",
    ]
    native = NativeNames(
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


def test_iteration_survives_internal_name_inserted_during_reference_resolution():
    native = NativeNames(
        [{"name": f"Range{i}", "formula": "=#REF!"} for i in range(1, 5)]
    )
    names = Names(native)
    seen = []
    for name in names:
        seen.append(name.name)
        if name.name == "Range2":
            with pytest.raises(ValueError, match="Broken reference"):
                name.refers_to_range
            # A lazy handle must still refer to Range2 after the index shift.
            assert name.name == "Range2"
    assert seen == ["Range1", "Range2", "Range3", "Range4"]
    assert [name.name for name in names] == seen
    assert len(names) == 4
    assert native.api[0]["name"] == "_xlfn.ANCHORARRAY"


def test_filtered_indices_edit_and_delete_the_intended_name():
    native = NativeNames(
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
