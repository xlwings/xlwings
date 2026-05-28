"""Sphinx extension to parse docstrings as MyST Markdown instead of RST.

Based on the approach proposed in:
https://github.com/executablebooks/MyST-Parser/issues/228#issuecomment-2584281870

Napoleon converts NumPy-style sections into RST field lists (:param:, :type:, etc.)
before our hook runs. We convert those field lists into MyST-friendly Markdown,
then wrap the whole docstring in an eval-myst directive.
"""

import re
from typing import Any

from myst_parser.parsers.sphinx_ import MystParser
from sphinx.application import Sphinx
from sphinx.util.docutils import SphinxDirective


class EvalMystDirective(SphinxDirective):
    has_content = True

    def run(self):
        document = self.state.document
        prev_children = document.children
        children = document.children = []
        try:
            parser = MystParser()
            parser.parse("\n".join(self.content), document)
        finally:
            document.children = prev_children
        return children


def _rst_inline_to_md(text: str) -> str:
    """Convert RST double backticks to MD single backticks."""
    return re.sub(r"``(.*?)``", r"`\1`", text)


def _collect_continuation(lines: list[str], i: int) -> tuple[list[str], int]:
    """Collect continuation lines (indented lines following a field).

    Handles blank lines within a field description — a blank line followed
    by more indented text is still part of the same field.
    """
    cont = []
    while i < len(lines):
        if lines[i] and lines[i][0] == " ":
            cont.append(lines[i].strip())
            i += 1
        elif lines[i] == "" and i + 1 < len(lines) and lines[i + 1].startswith(" "):
            # Blank line followed by indented text — still part of this field
            cont.append("")
            i += 1
        else:
            break
    return cont, i


def _convert_field_lists(lines: list[str]) -> list[str]:
    """Convert RST field lists from napoleon into MyST-friendly format."""
    result = []
    i = 0
    in_params = False

    while i < len(lines):
        line = lines[i]

        # Match :param name: description
        param_match = re.match(r"^:param (\w+):\s*(.*)", line)
        if param_match:
            name = param_match.group(1)
            desc_parts = [param_match.group(2)]
            i += 1
            cont, i = _collect_continuation(lines, i)
            desc_parts.extend(cont)

            # Check for :type name:
            type_str = ""
            if i < len(lines):
                type_match = re.match(
                    rf"^:type {re.escape(name)}:\s*(.*)", lines[i]
                )
                if type_match:
                    type_str = type_match.group(1)
                    i += 1

            if not in_params:
                result.append("**Parameters:**")
                result.append("")
                in_params = True

            desc = _rst_inline_to_md(" ".join(desc_parts))
            if type_str:
                result.append(f"* **{name}** (*{type_str}*) -- {desc}")
            else:
                result.append(f"* **{name}** -- {desc}")
            continue

        # Match :returns: or :return:
        ret_match = re.match(r"^:returns?:\s*(.*)", line)
        if ret_match:
            desc_parts = [ret_match.group(1)]
            i += 1
            cont, i = _collect_continuation(lines, i)
            desc_parts.extend(cont)

            type_str = ""
            if i < len(lines):
                rtype_match = re.match(r"^:rtype:\s*(.*)", lines[i])
                if rtype_match:
                    type_str = rtype_match.group(1)
                    i += 1

            if in_params:
                result.append("")
                in_params = False

            desc = _rst_inline_to_md(" ".join(desc_parts))
            if type_str:
                result.append(f"**Returns:** *{type_str}* -- {desc}")
            else:
                result.append(f"**Returns:** {desc}")
            result.append("")
            continue

        # Match standalone :rtype:
        rtype_match = re.match(r"^:rtype:\s*(.*)", line)
        if rtype_match:
            if in_params:
                result.append("")
                in_params = False
            result.append(f"**Return type:** *{rtype_match.group(1)}*")
            result.append("")
            i += 1
            continue

        # Match :raises ExcType:
        raises_match = re.match(r"^:raises (\w+):\s*(.*)", line)
        if raises_match:
            if in_params:
                result.append("")
                in_params = False
            result.append(f"**Raises:** **{raises_match.group(1)}** -- {raises_match.group(2)}")
            result.append("")
            i += 1
            continue

        # Match .. rubric:: heading (napoleon uses this for Examples)
        rubric_match = re.match(r"^\.\.\ rubric::\s*(.*)", line)
        if rubric_match:
            if in_params:
                result.append("")
                in_params = False
            result.append(f"**{rubric_match.group(1)}**")
            result.append("")
            i += 1
            continue

        # Match .. versionadded/changed/deprecated
        ver_match = re.match(
            r"^\.\.\s+(versionadded|versionchanged|deprecated)::\s*(.*)", line
        )
        if ver_match:
            if in_params:
                result.append("")
                in_params = False
            kind = ver_match.group(1)
            version = ver_match.group(2)
            if kind == "versionadded":
                result.append(f"*New in version {version}.*")
            elif kind == "versionchanged":
                result.append(f"*Changed in version {version}.*")
            elif kind == "deprecated":
                result.append(f"*Deprecated since version {version}.*")
            result.append("")
            i += 1
            continue

        # Match .. note::
        note_match = re.match(r"^\.\.\s+note::\s*(.*)", line)
        if note_match:
            if in_params:
                result.append("")
                in_params = False
            inline_text = note_match.group(1)
            if inline_text:
                # Inline note
                desc_parts = [inline_text]
                i += 1
                cont, i = _collect_continuation(lines, i)
                desc_parts.extend(cont)
                result.append(
                    "**Note:** " + _rst_inline_to_md(" ".join(desc_parts))
                )
            else:
                # Block note - collect indented content
                i += 1
                cont, i = _collect_continuation(lines, i)
                note_text = _rst_inline_to_md(" ".join(cont))
                result.append(f"**Note:** {note_text}")
            result.append("")
            continue

        # Match .. code-block:: lang
        code_match = re.match(r"^\.\.\s+code-block::\s*(.*)", line)
        if code_match:
            lang = code_match.group(1)
            i += 1
            # Skip blank line after directive
            if i < len(lines) and lines[i].strip() == "":
                i += 1
            # Collect indented code
            code_lines = []
            while i < len(lines) and (lines[i] == "" or lines[i].startswith("    ")):
                if lines[i] == "":
                    code_lines.append("")
                else:
                    code_lines.append(lines[i][4:])  # remove 4-space indent
                i += 1
            # Trim trailing blank lines
            while code_lines and code_lines[-1] == "":
                code_lines.pop()
            result.append(f"```{lang}")
            result.extend(code_lines)
            result.append("```")
            result.append("")
            continue

        # Convert remaining RST inline literals
        converted = _rst_inline_to_md(line)

        if in_params and line.strip() == "":
            in_params = False

        result.append(converted)
        i += 1

    return result


def process_docstring(
    app: Sphinx,
    what: str,
    name: str,
    obj: Any,
    options: dict[str, bool],
    lines: list[str],
):
    converted = _convert_field_lists(lines)
    lines[:] = [".. eval-myst::", ""] + ["    " + line for line in converted]


def setup(app: Sphinx):
    app.add_directive("eval-myst", EvalMystDirective)
    # Priority 600 ensures this runs after napoleon (default 500)
    app.connect("autodoc-process-docstring", process_docstring, priority=600)
    return {"version": "0.1.0", "parallel_read_safe": True}
