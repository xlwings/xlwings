"""
Required Notice: Copyright (C) Zoomer Analytics GmbH.

xlwings PRO is dual-licensed under one of the following licenses:

* PolyForm Noncommercial License 1.0.0 (for noncommercial use):
  https://polyformproject.org/licenses/noncommercial/1.0.0
* xlwings PRO License (for commercial use):
  https://github.com/xlwings/xlwings/blob/main/LICENSE_PRO.txt

Commercial licenses can be purchased at https://www.xlwings.org
"""

import sys
import warnings

try:
    import mistune
except ImportError:
    mistune = None
from ...conversion import Converter


class Style:
    def __init__(self, display_name=None):
        if display_name:
            self.display_name = display_name
        else:
            self.display_name = ""

    def __repr__(self):
        s = ""
        for attribute in vars(self):
            if getattr(self, attribute) and attribute != "display_name":
                s += f"{self.display_name}.{attribute}: {getattr(self, attribute)}\n"
        return s.replace("\n\n", "\n")


class FontStyle(Style):
    def __init__(
        self,
        display_name=None,
        color=None,
        size=None,
        bold=None,
        italic=None,
        name=None,
    ):
        super().__init__(display_name=display_name)
        self.color = color
        self.size = size
        self.bold = bold
        self.italic = italic
        self.name = name


class MarkdownStyle:
    """
    ``MarkdownStyle`` defines how ``Markdown`` objects are being rendered in Excel cells
    or shapes. Start by instantiating a ``MarkdownStyle`` object. Printing it will show
    you the current (default) style:

    >>> style = MarkdownStyle()
    >>> style
    <MarkdownStyle>
    h1.font: .bold: True
    h1.blank_lines_after: 1
    paragraph.blank_lines_after: 1
    unordered_list.bullet_character: •
    unordered_list.blank_lines_after: 1
    strong.bold: True
    emphasis.italic: True

    You can override the defaults, e.g., to make ``**strong**`` text red instead of
    bold, do this:

    >>> style.strong.bold = False
    >>> style.strong.color = (255, 0, 0)
    >>> style.strong
    strong.color: (255, 0, 0)

    .. versionadded:: 0.23.0
    """

    class __Heading1(Style):
        def __init__(self):
            super().__init__(display_name="h1")
            self.font = FontStyle(bold=True)
            self.blank_lines_after = 0

    class __Paragraph(Style):
        def __init__(self):
            super().__init__(display_name="paragraph")
            self.blank_lines_after = 1

    class __UnorderedList(Style):
        def __init__(self):
            super().__init__(display_name="unordered_list")
            self.bullet_character = "\u2022"
            self.blank_lines_after = 1

    def __init__(self):
        self.h1 = self.__Heading1()
        self.paragraph = self.__Paragraph()
        self.unordered_list = self.__UnorderedList()
        self.strong = FontStyle(display_name="strong", bold=True)
        self.emphasis = FontStyle(display_name="emphasis", italic=True)

    def __repr__(self):
        s = "<MarkdownStyle>\n"
        for attribute in vars(self):
            s += f"{getattr(self, attribute)}"
        return s


class Markdown:
    """
    Markdown objects can be assigned to a single cell or shape via ``myrange.value`` or
    ``myshape.text``. They accept a string in Markdown format which will cause the text
    in the cell to be formatted accordingly. They can also be used in
    ``mysheet.render_template()``.

    .. note:: On macOS, formatting is currently not supported, but things like bullet
              points will still work.

    Arguments
    ---------
    text : str
        The text in Markdown syntax

    style : MarkdownStyle object, optional
        The MarkdownStyle object defines how the text will be formatted.

    Examples
    --------

    >>> mysheet['A1'].value = Markdown("A text with *emphasis* and **strong** style.")
    >>> myshape.text = Markdown("A text with *emphasis* and **strong** style.")

    .. versionadded:: 0.23.0
    """

    def __init__(self, text, style=MarkdownStyle()):
        self.text = text
        self.style = style


class MarkdownConverter(Converter):
    @classmethod
    def write_value(cls, value, options):
        return render_text(value.text, value.style)


def traverse_ast_node(tree, data=None, level=0):
    data = (
        {
            "length": [],
            "type": [],
            "parent_type": [],
            "text": [],
            "parents": [],
            "level": [],
        }
        if data is None
        else data
    )
    for element in tree:
        data["parents"] = data["parents"][:level]
        if "children" in element:
            data["parents"].append(element)
            traverse_ast_node(element["children"], data, level=level + 1)
        else:
            data["level"].append(level)
            data["parent_type"].append([parent["type"] for parent in data["parents"]])
            data["type"].append(element["type"])
            if element["type"] == "text":
                marker = "text" if mistune.__version__.startswith("2") else "raw"
                data["length"].append(len(element[marker]))
                data["text"].append(element[marker])
            elif element["type"] in ("linebreak", "softbreak"):
                # mistune v2 uses linebreak, mistune v3 uses softbreak
                data["length"].append(1)
                data["text"].append("\n")
    return data


def flatten_ast(value):
    if not mistune:
        raise ImportError(
            "For xlwings Reports, "
            "you need to install mistune via 'pip/conda install mistune'"
        )
    if mistune.__version__.startswith("0"):
        raise ImportError(
            "Only mistune v2.x and v3.x are supported. "
            f"You have version {mistune.__version__}"
        )
    elif mistune.__version__.startswith("2"):
        parse_ast = mistune.create_markdown(renderer=mistune.AstRenderer())
    else:
        parse_ast = mistune.create_markdown(renderer="ast")
    ast = parse_ast(value)
    flat_ast = []
    for node in ast:
        rv = traverse_ast_node([node])
        del rv["parents"]
        flat_ast.append(rv)
    return flat_ast


def render_text(text, style):
    flat_ast = flatten_ast(text)
    output = ""
    for node in flat_ast:
        # heading/list currently don't respect the level
        if "heading" in node["parent_type"][0]:
            output += "".join(node["text"])
            output += "\n" + style.h1.blank_lines_after * "\n"
        elif "paragraph" in node["parent_type"][0]:
            output += "".join(node["text"])
            output += "\n" + style.paragraph.blank_lines_after * "\n"
        elif "list" in node["parent_type"][0]:
            for j in node["text"]:
                output += f"{style.unordered_list.bullet_character} {j}\n"
            output += style.unordered_list.blank_lines_after * "\n"
    return output.rstrip("\n")


def format_text(parent, text, style):
    if sys.platform.startswith("darwin"):
        # Characters formatting is broken because of a bug in AppleScript/Excel 2016
        warnings.warn("Markdown formatting is currently ignored on macOS.")
        return
    flat_ast = flatten_ast(text)
    position = 0
    for node in flat_ast:
        if "heading" in node["parent_type"][0]:
            node_length = sum(node["length"]) + style.h1.blank_lines_after + 1
            apply_style_to_font(
                style.h1.font, parent.characters[position : position + node_length].font
            )
        elif "paragraph" in node["parent_type"][0]:
            node_length = sum(node["length"]) + style.paragraph.blank_lines_after + 1
            intra_node_position = position
            for ix, j in enumerate(node["parent_type"]):
                selection = slice(
                    intra_node_position, intra_node_position + node["length"][ix]
                )
                if "strong" in j:
                    apply_style_to_font(style.strong, parent.characters[selection].font)
                elif "emphasis" in j:
                    apply_style_to_font(
                        style.emphasis, parent.characters[selection].font
                    )
                intra_node_position += node["length"][ix]
        elif "list" in node["parent_type"][0]:
            node_length = sum(node["length"]) + style.unordered_list.blank_lines_after
            for _ in node["text"]:
                # TODO: check ast level to allow nested **strong** etc.
                node_length += 3  # bullet, space and new line
        else:
            node_length = sum(node["length"])
        position += node_length


def apply_style_to_font(style_object, font_object):
    for attribute in vars(style_object):
        if getattr(style_object, attribute):
            setattr(font_object, attribute, getattr(style_object, attribute))
