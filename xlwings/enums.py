"""String enums for the public API.

Plain lowercase strings remain the documented way to pass these values;
the enums exist for autocomplete and typo-catching. Being `StrEnum` members,
they compare, hash and serialize exactly like their string values.
"""

from enum import StrEnum


class BorderIndex(StrEnum):
    """The eight individual sides of `Range.borders`.

    Group aliases (`"outside"`, `"inside"`, `"all"`, `"everything"`) are
    accepted by `Borders.set()`/`Borders.clear()` as plain strings only.

    ```{versionadded} 0.37.1
    ```
    """

    edge_top = "edge_top"
    edge_bottom = "edge_bottom"
    edge_left = "edge_left"
    edge_right = "edge_right"
    inside_vertical = "inside_vertical"
    inside_horizontal = "inside_horizontal"
    diagonal_down = "diagonal_down"
    diagonal_up = "diagonal_up"


class BorderLineStyle(StrEnum):
    """Line styles for `Border.line_style`. `none` removes the border.

    ```{versionadded} 0.37.1
    ```
    """

    continuous = "continuous"
    dash = "dash"
    dash_dot = "dash_dot"
    dash_dot_dot = "dash_dot_dot"
    dot = "dot"
    double = "double"
    slant_dash_dot = "slant_dash_dot"
    none = "none"


class BorderWeight(StrEnum):
    """Line weights for `Border.weight`.

    ```{versionadded} 0.37.1
    ```
    """

    hairline = "hairline"
    thin = "thin"
    medium = "medium"
    thick = "thick"
