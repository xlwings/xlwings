import asyncio
from unittest import mock

import pytest

import xlwings as xw
from xlwings import XlwingsError
from xlwings.pro import _xlremote as R


def _book():
    payload = {
        "client": "Office.js",
        "version": xw.__version__,
        "book": {"name": "B", "active_sheet_index": 0, "selection": "A1"},
        "names": [],
        "sheets": [
            {
                "name": "Sheet1",
                "values": [[1, 2], [3, 4]],
                "pictures": [],
                "tables": [],
            }
        ],
    }
    impl = R.App(R.Apps(), add_book=False).books.open(payload)
    return xw.Book(impl=impl)


def _entries():
    return [
        {
            "type": "CellValue",
            "stop_if_true": True,
            "operator": "LessThan",
            "formula1": "60",
            "formula2": None,
            "fill_color": "#ffff00",
            "font_color": None,
            "font_bold": None,
            "font_italic": True,
        },
        {"type": "DataBar", "stop_if_true": None},
        {"type": "PresetCriteria", "stop_if_true": False},
    ]


def test_public_types_are_exported():
    assert xw.ConditionalFormat.__name__ == "ConditionalFormat"
    assert xw.ConditionalFormats.__name__ == "ConditionalFormats"


def test_unloaded_collection_can_clear_without_a_read():
    book = _book()
    formats = book.sheets[0]["A1:B2"].conditional_formats

    formats.clear()

    action = book.impl.json()["actions"][-1]
    assert action["func"] == "clearConditionalFormats"
    assert action["args"] == []
    assert (action["row_count"], action["column_count"]) == (2, 2)


def test_unloaded_collection_points_inspection_at_async_getter():
    formats = _book().sheets[0]["A1"].conditional_formats
    with pytest.raises(NotImplementedError, match=r"get_conditional_formats\(\)"):
        len(formats)


def test_async_getter_is_lite_only_off_emscripten():
    with pytest.raises(NotImplementedError, match="only supported in xlwings Lite"):
        asyncio.run(_book().sheets[0]["A1"].get_conditional_formats())


def test_async_getter_preserves_order_and_unknown_rules():
    rng = _book().sheets[0]["A1:B2"]

    async def fake(self, key, method=None):
        assert key == "conditional_formats"
        return _entries()

    with mock.patch.object(R.Range, "_get_range_data", fake):
        formats = asyncio.run(rng.get_conditional_formats())

    assert len(formats) == 3
    assert [rule.type for rule in formats] == ["cell_value", "data_bar", "unknown"]
    assert [rule.stop_if_true for rule in formats] == [True, None, False]
    assert formats[0].type == "cell_value"
    assert formats(2).type == "data_bar"
    assert formats[0].operator == "less_than"
    assert formats[0].formula1 == "60"
    assert formats[0].formula2 is None
    assert formats[0].formula is None
    assert formats[0].fill_color == (255, 255, 0)
    assert formats[0].font_italic is True


def test_async_getter_ignores_inactive_formula2():
    rng = _book().sheets[0]["A1"]
    entry = {**_entries()[0], "formula2": "999"}

    async def fake(self, key, method=None):
        return [entry]

    with mock.patch.object(R.Range, "_get_range_data", fake):
        rule = asyncio.run(rng.get_conditional_formats())[0]

    assert rule.operator == "less_than"
    assert rule.formula2 is None

    entry["operator"] = "Between"
    with mock.patch.object(R.Range, "_get_range_data", fake):
        rule = asyncio.run(rng.get_conditional_formats())[0]
    assert rule.operator == "between"
    assert rule.formula2 == "999"


def test_multiple_deletes_keep_snapshot_positions_aligned():
    book = _book()
    rng = book.sheets[0]["A1:B2"]

    async def fake(self, key, method=None):
        return _entries()

    with mock.patch.object(R.Range, "_get_range_data", fake):
        formats = asyncio.run(rng.get_conditional_formats())

    rules = list(formats)
    rules[0].delete()
    rules[1].delete()

    actions = book.impl.json()["actions"]
    assert [action["func"] for action in actions] == [
        "deleteConditionalFormat",
        "deleteConditionalFormat",
    ]
    assert [action["args"][0] for action in actions] == [0, 0]
    assert actions[0]["args"][1] == _entries()[0]
    assert actions[1]["args"][1] == _entries()[1]
    assert [rule.type for rule in formats] == ["unknown"]


def test_deleted_rule_cannot_be_deleted_twice():
    rng = _book().sheets[0]["A1"]

    async def fake(self, key, method=None):
        return _entries()[:1]

    with mock.patch.object(R.Range, "_get_range_data", fake):
        rule = asyncio.run(rng.get_conditional_formats())[0]
    rule.delete()
    with pytest.raises(XlwingsError, match="no longer in its collection"):
        rule.delete()


def test_delete_queues_complete_normalized_visual_snapshot():
    book = _book()
    rng = book.sheets[0]["A1:A10"]
    entry = {
        "type": "DataBar",
        "stop_if_true": None,
        "bar_color": "#638ec6",
        "gradient": True,
        "show_value": True,
        "threshold_types": ["Automatic", "Number"],
        "thresholds": [None, 100],
    }

    async def fake(self, key, method=None):
        return [entry]

    with mock.patch.object(R.Range, "_get_range_data", fake):
        formats = asyncio.run(rng.get_conditional_formats())
    formats[0].delete()

    assert book.impl.json()["actions"][-1]["args"] == [
        0,
        {**entry, "thresholds": [None, "100"]},
    ]


def test_clear_updates_a_loaded_snapshot():
    rng = _book().sheets[0]["A1"]

    async def fake(self, key, method=None):
        return _entries()

    with mock.patch.object(R.Range, "_get_range_data", fake):
        formats = asyncio.run(rng.get_conditional_formats())
    formats.clear()
    assert len(formats) == 0


def test_add_cell_value_queues_complete_rule_and_returns_local_rule():
    book = _book()
    formats = book.sheets[0]["B2:B12"].conditional_formats

    rule = formats.add_cell_value(
        "less_than",
        60,
        fill_color="#FFFF00",
        font_italic=True,
    )

    action = book.impl.json()["actions"][-1]
    assert action["func"] == "addConditionalFormat"
    assert action["args"] == [
        {
            "type": "CellValue",
            "operator": "LessThan",
            "formula1": "60",
            "formula2": None,
            "stop_if_true": False,
            "fill_color": "#ffff00",
            "font_color": None,
            "font_bold": None,
            "font_italic": True,
        }
    ]
    assert rule.type == "cell_value"
    assert rule.operator == "less_than"
    assert rule.formula1 == "60"
    assert rule.fill_color == (255, 255, 0)


def test_add_custom_and_set_queue_ordered_mutations_and_update_local_state():
    book = _book()
    formats = book.sheets[0]["A2:D20"].conditional_formats
    rule = formats.add_custom("=$A2<>$B2", fill_color=(255, 242, 204))

    rule.set(formula="=$A2=$B2", font_bold=True, stop_if_true=True)

    actions = book.impl.json()["actions"]
    assert [action["func"] for action in actions] == [
        "addConditionalFormat",
        "setConditionalFormat",
    ]
    assert actions[1]["args"][0] == 0
    assert actions[1]["args"][2] == {
        "formula": "=$A2=$B2",
        "font_bold": True,
        "stop_if_true": True,
    }
    assert rule.formula == "=$A2=$B2"
    assert rule.font_bold is True
    assert rule.stop_if_true is True


def test_pending_adds_keep_returned_rule_positions_aligned():
    book = _book()
    formats = book.sheets[0]["A1:A10"].conditional_formats
    first = formats.add_cell_value("less_than", 10)
    formats.add_custom("=A1=0")

    first.set(formula1=20)

    action = book.impl.json()["actions"][-1]
    assert action["func"] == "setConditionalFormat"
    assert action["args"][0] == 1


def test_set_preserves_unspecified_cell_value_attributes():
    book = _book()
    rng = book.sheets[0]["B2:B12"]

    async def fake(self, key, method=None):
        return _entries()[:1]

    with mock.patch.object(R.Range, "_get_range_data", fake):
        rule = asyncio.run(rng.get_conditional_formats())[0]
    rule.set(formula1=70)

    action = book.impl.json()["actions"][-1]
    assert action["func"] == "setConditionalFormat"
    assert action["args"][1] == {
        "type": "CellValue",
        "stop_if_true": True,
        "operator": "LessThan",
        "formula1": "60",
        "formula2": None,
        "fill_color": "#ffff00",
        "font_color": None,
        "font_bold": None,
        "font_italic": True,
    }
    assert action["args"][2] == {"formula1": "70"}
    assert rule.operator == "less_than"
    assert rule.formula1 == "70"
    assert rule.fill_color == (255, 255, 0)


def test_set_can_switch_from_between_and_clears_formula2():
    book = _book()
    rng = book.sheets[0]["B2:B12"]
    entry = {
        **_entries()[0],
        "operator": "Between",
        "formula1": "60",
        "formula2": "80",
    }

    async def fake(self, key, method=None):
        return [entry]

    with mock.patch.object(R.Range, "_get_range_data", fake):
        rule = asyncio.run(rng.get_conditional_formats())[0]
    rule.set(operator="less_than")

    assert book.impl.json()["actions"][-1]["args"][2] == {
        "operator": "LessThan",
        "formula2": None,
    }
    assert rule.formula2 is None


@pytest.mark.parametrize(
    ("operator", "formula2", "message"),
    [
        ("between", None, "formula2 is required"),
        ("less_than", 80, "formula2 is only valid"),
        ("nonsense", None, "Invalid conditional-format operator"),
    ],
)
def test_add_cell_value_validates_operator_formula_pair(operator, formula2, message):
    formats = _book().sheets[0]["A1"].conditional_formats
    with pytest.raises(ValueError, match=message):
        formats.add_cell_value(operator, 60, formula2)


def test_custom_formula_must_start_with_equals():
    formats = _book().sheets[0]["A1"].conditional_formats
    with pytest.raises(ValueError, match="starting with"):
        formats.add_custom("A1>0")


def test_unknown_rule_cannot_be_edited():
    rng = _book().sheets[0]["A1"]

    async def fake(self, key, method=None):
        return _entries()[2:]

    with mock.patch.object(R.Range, "_get_range_data", fake):
        rule = asyncio.run(rng.get_conditional_formats())[0]
    with pytest.raises(NotImplementedError, match="unknown"):
        rule.set(stop_if_true=True)


def test_add_color_scale_with_custom_number_thresholds():
    book = _book()
    formats = book.sheets[0]["B2:B20"].conditional_formats

    rule = formats.add_color_scale(
        ["#f8696b", "#ffeb84", "#63be7b"],
        thresholds=[0, 50, 100],
        threshold_type="number",
    )

    action = book.impl.json()["actions"][-1]
    assert action["func"] == "addConditionalFormat"
    assert action["args"] == [
        {
            "type": "ColorScale",
            "stop_if_true": None,
            "fill_color": None,
            "font_color": None,
            "font_bold": None,
            "font_italic": None,
            "colors": ["#f8696b", "#ffeb84", "#63be7b"],
            "threshold_types": ["Number", "Number", "Number"],
            "thresholds": [0, 50, 100],
        }
    ]
    assert rule.colors == ((248, 105, 107), (255, 235, 132), (99, 190, 123))
    assert rule.threshold_types == ("number", "number", "number")
    assert rule.thresholds == (0, 50, 100)


def test_add_color_scale_uses_distribution_defaults():
    rule = (
        _book()
        .sheets[0]["A1:A10"]
        .conditional_formats.add_color_scale(["#ff0000", "#ffff00", "#00ff00"])
    )

    assert rule.threshold_types == (
        "lowest_value",
        "percentile",
        "highest_value",
    )
    assert rule.thresholds == (None, 50, None)


def test_add_data_bar_with_one_automatic_bound():
    book = _book()
    rule = book.sheets[0]["A1:A10"].conditional_formats.add_data_bar(
        "#638ec6",
        maximum=100,
        gradient=False,
        show_value=False,
    )

    action = book.impl.json()["actions"][-1]
    assert action["args"][0]["type"] == "DataBar"
    assert action["args"][0]["bar_color"] == "#638ec6"
    assert action["args"][0]["threshold_types"] == ["Automatic", "Number"]
    assert action["args"][0]["thresholds"] == [None, 100]
    assert rule.bar_color == (99, 142, 198)
    assert rule.gradient is False
    assert rule.show_value is False


def test_add_icon_set_with_custom_thresholds():
    book = _book()
    rule = book.sheets[0]["A1:A10"].conditional_formats.add_icon_set(
        "3_traffic_lights_1",
        thresholds=[60, 80],
        show_value=False,
        reverse_order=True,
    )

    entry = book.impl.json()["actions"][-1]["args"][0]
    assert entry["type"] == "IconSet"
    assert entry["icon_set"] == "ThreeTrafficLights1"
    assert entry["threshold_types"] == ["Number", "Number"]
    assert entry["thresholds"] == [60, 80]
    assert rule.icon_set == "3_traffic_lights_1"
    assert rule.show_value is False
    assert rule.reverse_order is True


def test_add_icon_set_uses_equal_percent_bands():
    rule = _book().sheets[0]["A1:A10"].conditional_formats.add_icon_set("5_quarters")

    assert rule.threshold_types == ("percent",) * 4
    assert rule.thresholds == (20, 40, 60, 80)


@pytest.mark.parametrize(
    ("call", "message"),
    [
        (lambda formats: formats.add_color_scale(["#ff0000"]), "two or three"),
        (
            lambda formats: formats.add_color_scale(
                ["#ff0000", "#00ff00"], thresholds=[0]
            ),
            "exactly 2",
        ),
        (
            lambda formats: formats.add_icon_set("3_arrows", thresholds=[80, 60]),
            "strictly increasing",
        ),
        (
            lambda formats: formats.add_icon_set("rainbows"),
            "Invalid conditional-format icon set",
        ),
        (
            lambda formats: formats.add_data_bar("blue", minimum=100, maximum=0),
            "minimum must be less",
        ),
        (
            lambda formats: formats.add_color_scale(
                ["#ff0000", "#00ff00"],
                thresholds=[0, 101],
                threshold_type="percent",
            ),
            "between 0 and 100",
        ),
    ],
)
def test_visual_rule_validation(call, message):
    with pytest.raises((TypeError, ValueError), match=message):
        call(_book().sheets[0]["A1"].conditional_formats)


def test_async_getter_exposes_visual_rule_details():
    rng = _book().sheets[0]["A1:A10"]
    entries = [
        {
            "type": "ColorScale",
            "stop_if_true": None,
            "colors": ["#f8696b", "#ffeb84", "#63be7b"],
            "threshold_types": ["LowestValue", "Percentile", "HighestValue"],
            "thresholds": [None, 50, None],
        },
        {
            "type": "DataBar",
            "stop_if_true": None,
            "bar_color": "#638ec6",
            "gradient": True,
            "show_value": True,
            "threshold_types": ["Automatic", "Number"],
            "thresholds": [None, 100],
        },
        {
            "type": "IconSet",
            "stop_if_true": None,
            "icon_set": "ThreeTrafficLights1",
            "show_value": False,
            "reverse_order": True,
            "threshold_types": ["Number", "Number"],
            "thresholds": [60, 80],
        },
    ]

    async def fake(self, key, method=None):
        return entries

    with mock.patch.object(R.Range, "_get_range_data", fake):
        formats = asyncio.run(rng.get_conditional_formats())

    assert formats[0].colors[1] == (255, 235, 132)
    assert formats[1].bar_color == (99, 142, 198)
    assert formats[1].threshold_types == ("automatic", "number")
    assert formats[2].icon_set == "3_traffic_lights_1"
    assert formats[2].thresholds == (60, 80)
