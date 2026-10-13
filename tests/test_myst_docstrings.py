"""Keyword argument documentation renders through the MyST docstring hook."""

from pathlib import Path

import pytest

Sphinx = pytest.importorskip("sphinx.application").Sphinx


def test_keyword_arguments_render_as_fields_with_links(tmp_path, monkeypatch):
    source = tmp_path / "source"
    source.mkdir()
    extension = Path(__file__).resolve().parents[1] / "docs" / "_ext"
    monkeypatch.syspath_prepend(str(extension))
    monkeypatch.syspath_prepend(str(source))
    (source / "conf.py").write_text(
        'extensions = ["myst_parser", "sphinx.ext.autodoc", '
        '"sphinx.ext.napoleon", "myst_docstrings"]\n'
        'autodoc_typehints = "description"\n'
        'master_doc = "index"\n'
    )
    (source / "docstring_example.py").write_text(
        "from typing import TypeVar\n\n"
        "_Result = TypeVar('_Result')\n\n"
        "def options():\n"
        '    """Configure conversion."""\n'
        "    pass\n\n"
        "def read(value: _Result, **kwargs) -> _Result:\n"
        '    """Read values.\n\n'
        "    Args:\n"
        "        value (callable): The function to configure.\n\n"
        "    Keyword Args:\n"
        "        ndim: Use `2` for rows; see "
        "{func}`options <docstring_example.options>`.\n"
        "\n    Returns:\n"
        "        callable: The configured function.\n"
        '    """\n'
        "    pass\n"
    )
    (source / "index.md").write_text(
        "# Example\n\n```{eval-rst}\n"
        ".. automodule:: docstring_example\n    :members:\n```\n"
    )
    output = tmp_path / "html"
    from sphinx.util.docutils import docutils_namespace

    with docutils_namespace():
        app = Sphinx(
            str(source),
            str(source),
            str(output),
            str(tmp_path / "doctrees"),
            "html",
            warningiserror=True,
            freshenv=True,
        )
        app.build(force_all=True)

    assert app.statuscode == 0
    html = (output / "index.html").read_text()
    assert "Keyword Arguments" in html
    assert ":keyword ndim:" not in html
    assert 'href="#docstring_example.options"' in html
    assert '<span class="pre">2</span>' in html
    assert html.count("Return type") == 1
    assert "_Result" not in html
    assert "The configured function." in html
