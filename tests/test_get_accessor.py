from xlwings.conversion import (
    _get_accessor,
    ValueAccessor,
    RawValueAccessor,
    RangeAccessor,
)


def test_none_default() -> None:
    """
    When `convert` is `None`, the default Accessor class is returned.
    """
    assert ValueAccessor == _get_accessor(
        convert=None,
        default=ValueAccessor,
        registered={
            "raw": RawValueAccessor,
            "range": RangeAccessor,
        },
    )


def test_registered_name() -> None:
    """
    `convert` may be a registered name that maps to an Accessor class.
    """
    assert ValueAccessor == _get_accessor(
        convert="value",
        default=RawValueAccessor,
        registered={
            "range": RangeAccessor,
            "value": ValueAccessor,
        },
    )


def test_not_registered() -> None:
    """
    When `convert` is an unregistered name, the default Accessor class is returned.
    """
    assert ValueAccessor == _get_accessor(
        convert="foo",
        default=ValueAccessor,
        registered={
            "raw": RawValueAccessor,
            "range": RangeAccessor,
        },
    )


class CustomClass: ...


class CustomAccessor(ValueAccessor): ...


def test_custom_class() -> None:
    """
    `convert` may be a registered type that maps to an Accessor class.
    """
    assert CustomAccessor == _get_accessor(
        convert=CustomClass,
        default=RawValueAccessor,
        registered={
            CustomClass: CustomAccessor,
            "range": RangeAccessor,
            "value": ValueAccessor,
        },
    )


def test_custom_class_not_registered() -> None:
    """
    When `convert` is an unregistered type, the default Accessor class is returned.
    """
    assert ValueAccessor == _get_accessor(
        convert=CustomClass,
        default=ValueAccessor,
        registered={
            "raw": RawValueAccessor,
            "range": RangeAccessor,
        },
    )


def test_unregistered_none_type() -> None:
    """
    When `convert` is `NoneType`, and no special accessor was registered for it,
    the default Accessor class is returned.

    See https://github.com/xlwings/xlwings/issues/2666
    """
    assert ValueAccessor == _get_accessor(
        convert=type(None),
        default=ValueAccessor,
        registered={
            "raw": RawValueAccessor,
            "range": RangeAccessor,
        },
    )


def test_unregistered_iterable_type() -> None:
    """
    When `convert` is `Iterable`, and no special accessor was registered for it,
    the default Accessor class is returned.

    See https://github.com/xlwings/xlwings/issues/2666
    """
    from collections.abc import Iterable

    assert ValueAccessor == _get_accessor(
        convert=Iterable,
        default=ValueAccessor,
        registered={
            "raw": RawValueAccessor,
            "range": RangeAccessor,
        },
    )
