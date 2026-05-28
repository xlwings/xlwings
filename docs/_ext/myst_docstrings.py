"""Sphinx extension to parse docstrings as MyST Markdown instead of RST.

Based on the approach proposed in:
https://github.com/executablebooks/MyST-Parser/issues/228#issuecomment-2584281870

Napoleon converts NumPy-style sections into RST field lists (:param:, :type:, etc.)
before our hook runs. We keep those field lists as native RST (so Sphinx/Furo renders
them with proper styling), and only wrap the non-field-list prose in an eval-myst
directive so that Markdown formatting (backticks, code blocks, notes, etc.) is parsed.
"""

import re
from typing import Any

from myst_parser.parsers.sphinx_ import MystParser
from sphinx.application import Sphinx
from sphinx.util.docutils import SphinxDirective


class EvalMystDirective(SphinxDirective):
    has_content = True

    def run(self):
        from docutils import nodes

        document = self.state.document
        # Save state that MyST modifies
        prev_children = document.children
        prev_sub_defs = dict(document.substitution_defs)
        prev_sub_names = dict(document.substitution_names)
        children = document.children = []
        try:
            parser = MystParser()
            parser.parse("\n".join(self.content), document)
        finally:
            document.children = prev_children
            # Restore substitution state to prevent "Duplicate substitution
            # definition" errors when eval-myst runs multiple times per doc.
            document.substitution_defs = prev_sub_defs
            document.substitution_names = prev_sub_names
        # Filter out substitution_definition nodes
        return [
            c for c in children if not isinstance(c, nodes.substitution_definition)
        ]


def _rst_inline_to_md(text: str) -> str:
    """Convert RST double backticks to MD single backticks."""
    return re.sub(r"``(.*?)``", r"`\1`", text)


def _md_inline_to_rst(text: str) -> str:
    """Convert MD single backticks to RST double backticks.

    Used for content that stays as native RST (field lists).
    Avoids converting already-double backticks.
    """
    return re.sub(r"(?<!`)(`)(?!`)(.+?)(?<!`)(`)", r"``\2``", text)


def _is_rst_field(line: str) -> bool:
    """Check if a line is an RST field list entry."""
    return bool(re.match(r"^:(param|type|returns?|rtype|raises)\b", line))


def _is_rst_directive(line: str) -> bool:
    """Check if a line is an RST directive."""
    return bool(
        re.match(
            r"^\.\.\s+(versionadded|versionchanged|deprecated|seealso|note|warning|rubric|code-block)::",
            line,
        )
    )


def process_docstring(
    app: Sphinx,
    what: str,
    name: str,
    obj: Any,
    options: dict[str, bool],
    lines: list[str],
):
    """Process docstring: wrap Markdown prose in eval-myst, keep RST field lists as-is.

    We split the docstring into segments:
    - RST segments (field lists, directives) are kept as native RST
    - Prose segments are wrapped in .. eval-myst:: for MyST parsing
    """
    if not lines:
        return

    # Identify segments: each is either "rst" or "myst"
    segments = []  # list of (type, lines)
    current_type = "myst"
    current_lines = []

    i = 0
    while i < len(lines):
        line = lines[i]

        if _is_rst_field(line):
            # Switch to RST mode for field lists
            if current_type != "rst" and current_lines:
                segments.append((current_type, current_lines))
                current_lines = []
            current_type = "rst"
            # Docstrings use MD backticks, but field list content stays as
            # native RST, so convert single backticks to double.
            current_lines.append(_md_inline_to_rst(line))
            i += 1
            # Collect continuation lines (indented)
            while i < len(lines) and lines[i] and lines[i][0] == " ":
                current_lines.append(_md_inline_to_rst(lines[i]))
                i += 1
            continue

        if _is_rst_directive(line):
            # Keep RST directives as-is
            if current_type != "rst" and current_lines:
                segments.append((current_type, current_lines))
                current_lines = []
            current_type = "rst"
            current_lines.append(line)
            i += 1
            # Collect directive content (indented or blank followed by indented)
            while i < len(lines):
                if lines[i] and lines[i][0] == " ":
                    current_lines.append(lines[i])
                    i += 1
                elif (
                    lines[i] == ""
                    and i + 1 < len(lines)
                    and lines[i + 1].startswith(" ")
                ):
                    current_lines.append(lines[i])
                    i += 1
                else:
                    break
            continue

        # Regular prose line — MyST mode
        if current_type != "myst" and current_lines:
            segments.append((current_type, current_lines))
            current_lines = []
        current_type = "myst"
        current_lines.append(_rst_inline_to_md(line))
        i += 1

    if current_lines:
        segments.append((current_type, current_lines))

    # Build output: wrap MyST segments in eval-myst, keep RST as-is
    output = []
    for seg_type, seg_lines in segments:
        if seg_type == "myst":
            output.append(".. eval-myst::")
            output.append("")
            for sl in seg_lines:
                output.append("    " + sl)
            output.append("")
        else:
            output.extend(seg_lines)
            output.append("")

    lines[:] = output


def setup(app: Sphinx):
    app.add_directive("eval-myst", EvalMystDirective)
    # Priority 600 ensures this runs after napoleon (default 500)
    app.connect("autodoc-process-docstring", process_docstring, priority=600)
    return {"version": "0.1.0", "parallel_read_safe": True}
