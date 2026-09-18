"""Round-trip conditional-format tests against real desktop Excel.

These tests intentionally exercise the native Windows and macOS engines. The
remote-engine tests only verify serialized actions and therefore cannot catch
differences in the COM and AppleScript object models.
"""

import sys

import pytest

import xlwings as xw


@pytest.fixture(scope="module")
def app():
    with xw.App(visible=False) as app:
        yield app


@pytest.fixture
def rng(app):
    book = app.books.add()
    rng = book.sheets[0]["B2:B20"]
    rng.value = [[value] for value in range(19)]
    yield rng
    book.close()


def _without_equals(value):
    """Normalize Excel's platform-dependent leading equals sign."""
    return value.removeprefix("=")


def test_collection_starts_empty(rng):
    formats = rng.conditional_formats

    assert formats.api is not None
    assert len(formats) == 0
    assert list(formats) == []


def test_add_and_set_cell_value_rule(rng):
    rule = rng.conditional_formats.add_cell_value(
        "between",
        5,
        10,
        fill_color="#ffff00",
        font_color="#ff0000",
        font_bold=True,
        font_italic=True,
        stop_if_true=True,
    )

    assert rule.type == "cell_value"
    assert rule.operator == "between"
    assert _without_equals(rule.formula1) == "5"
    assert _without_equals(rule.formula2) == "10"
    assert rule.formula is None
    assert rule.fill_color == (255, 255, 0)
    assert rule.font_color == (255, 0, 0)
    assert rule.font_bold is True
    assert rule.font_italic is True
    assert rule.stop_if_true is True

    rule.set(
        operator="greater_than",
        formula1=12,
        fill_color="#00ff00",
        font_color="#0000ff",
        font_bold=False,
        font_italic=False,
        stop_if_true=False,
    )

    assert rule.operator == "greater_than"
    assert _without_equals(rule.formula1) == "12"
    assert rule.formula2 is None
    assert rule.fill_color == (0, 255, 0)
    assert rule.font_color == (0, 0, 255)
    assert rule.font_bold is False
    assert rule.font_italic is False
    assert rule.stop_if_true is False


def test_add_and_set_custom_rule(rng):
    rule = rng.conditional_formats.add_custom(
        "=B2>5", fill_color="#ffc7ce", stop_if_true=True
    )

    assert rule.type == "custom"
    assert rule.formula == "=B2>5"
    assert rule.operator is None
    assert rule.formula1 is None
    assert rule.formula2 is None
    assert rule.fill_color == (255, 199, 206)
    assert rule.stop_if_true is True

    rule.set(formula="=B2>10", font_bold=True, stop_if_true=False)

    assert rule.formula == "=B2>10"
    assert rule.font_bold is True
    assert rule.stop_if_true is False


def test_add_color_scale_with_custom_number_thresholds(rng):
    # Keep this in sync with the public documentation example: it caught an
    # AppleScript collection-addressing bug that action serialization missed.
    rule = rng.conditional_formats.add_color_scale(
        ["#f8696b", "#ffeb84", "#63be7b"],
        thresholds=[0, 50, 100],
        threshold_type="number",
    )

    assert rule.type == "color_scale"
    assert rule.colors == ((248, 105, 107), (255, 235, 132), (99, 190, 123))
    assert rule.threshold_types == ("number", "number", "number")
    assert rule.thresholds == (0, 50, 100)
    assert rule.stop_if_true is None


def test_add_color_scale_with_distribution_defaults(rng):
    rule = rng.conditional_formats.add_color_scale(["#ff0000", "#00ff00"])

    assert rule.type == "color_scale"
    assert rule.colors == ((255, 0, 0), (0, 255, 0))
    assert rule.threshold_types == ("lowest_value", "highest_value")
    assert rule.thresholds == (None, None)


def test_add_data_bar_with_automatic_and_custom_bounds(rng):
    rule = rng.conditional_formats.add_data_bar(
        "#638ec6",
        maximum=100,
        gradient=False,
        show_value=False,
    )

    assert rule.type == "data_bar"
    assert rule.bar_color == (99, 142, 198)
    assert rule.gradient is False
    assert rule.show_value is False
    assert rule.threshold_types == ("automatic", "number")
    assert rule.thresholds == (None, 100)
    assert rule.stop_if_true is None


def test_add_icon_set_with_custom_thresholds(rng):
    rule = rng.conditional_formats.add_icon_set(
        "3_traffic_lights_1",
        thresholds=[6, 12],
        show_value=False,
        reverse_order=True,
    )

    assert rule.type == "icon_set"
    assert rule.icon_set == "3_traffic_lights_1"
    assert rule.show_value is False
    assert rule.reverse_order is True
    assert rule.threshold_types == ("number", "number")
    assert rule.thresholds == (6, 12)
    assert rule.stop_if_true is None


def test_icon_set_uses_equal_percent_bands_by_default(rng):
    rule = rng.conditional_formats.add_icon_set("5_quarters")

    assert rule.icon_set == "5_quarters"
    assert rule.threshold_types == ("percent",) * 4
    assert rule.thresholds == (20, 40, 60, 80)


@pytest.mark.skipif(sys.platform != "darwin", reason="Requires the appscript engine")
def test_mac_icon_set_indexes_match_generated_terminology():
    from xlwings import _xlmac

    terminology = dict(_xlmac.mac_dict.enums)
    assert {
        name: int.from_bytes(terminology[keyword.AS_name][2:], "big")
        for name, keyword in _xlmac._CONDITIONAL_FORMAT_ICON_SET_TO_KW.items()
    } == _xlmac._CONDITIONAL_FORMAT_ICON_SET_INDEX


def test_clear_preserves_rule_outside_target_range(rng):
    rng.conditional_formats.add_cell_value("less_than", 10)
    cleared = rng.sheet["B2:B10"]
    remaining = rng.sheet["B11:B20"]

    cleared.conditional_formats.clear()

    assert len(cleared.conditional_formats) == 0
    assert [rule.type for rule in remaining.conditional_formats] == ["cell_value"]


def test_mixed_collection_order_delete_and_clear(rng):
    formats = rng.conditional_formats
    formats.add_cell_value("less_than", 5)
    formats.add_custom("=B2=10")
    formats.add_color_scale(["#ff0000", "#00ff00"])
    formats.add_data_bar("#638ec6")
    formats.add_icon_set("3_arrows")

    assert [rule.type for rule in formats] == [
        "icon_set",
        "data_bar",
        "color_scale",
        "custom",
        "cell_value",
    ]
    assert formats[0].icon_set == "3_arrows"
    assert formats[1].bar_color == (99, 142, 198)
    assert formats[2].colors == ((255, 0, 0), (0, 255, 0))
    assert formats[3].formula == "=B2=10"
    assert _without_equals(formats[4].formula1) == "5"

    formats[2].delete()
    assert [rule.type for rule in formats] == [
        "icon_set",
        "data_bar",
        "custom",
        "cell_value",
    ]

    formats.clear()
    assert len(formats) == 0
    assert list(formats) == []
