"""Overloaded API functions display one callable implementation signature."""

from pathlib import Path

import pytest

Sphinx = pytest.importorskip("sphinx.application").Sphinx


def test_overloads_render_one_signature_with_types_in_description(
    tmp_path, monkeypatch
):
    source = tmp_path / "source"
    source.mkdir()
    monkeypatch.syspath_prepend(
        str(Path(__file__).resolve().parents[1] / "docs" / "_ext")
    )
    monkeypatch.syspath_prepend(str(source))
    (source / "conf.py").write_text(
        'extensions = ["myst_parser", "sphinx.ext.autodoc", '
        '"sphinx.ext.napoleon", "myst_docstrings", "implementation_signatures"]\n'
        'autodoc_typehints = "description"\n'
        'master_doc = "index"\n'
    )
    (source / "signature_example.py").write_text(
        "from typing import overload\n\n"
        "@overload\ndef script(f: str) -> str: ...\n\n"
        "@overload\ndef script(f: None = None, *, name: str = '') -> str: ...\n\n"
        "def script(f: str | None = None, *, name: str = '', **kwargs):\n"
        '    """Register a script.\n\n'
        "    Args:\n"
        "        f: Function to decorate.\n"
        "        name: Display name.\n"
        '    """\n'
        "    return f\n"
    )
    (source / "index.md").write_text(
        "# Example\n\n```{eval-rst}\n"
        ".. autofunction:: signature_example.script\n```\n"
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
    tree = app.env.get_doctree("index")
    from sphinx import addnodes

    headings = list(tree.findall(addnodes.desc_signature))
    assert len(headings) == 1
    assert (
        headings[0].astext() == "signature_example.script(f=None, *, name='', **kwargs)"
    )
    assert "str" in tree.astext()
    assert 'id="signature_example.script"' in (output / "index.html").read_text()
