import os
import sys

sys.path.insert(0, os.path.abspath(".."))

# -- Handle unavailable packages/modules on build machine -----------------------

# pywin32 can't be installed on non-Windows OS (e.g. on Read-the-Docs), therefore mock it


class Mock(object):
    __all__ = []

    def __init__(self, *args, **kwargs):
        pass

    def __call__(self, *args, **kwargs):
        return Mock()

    @classmethod
    def __getattr__(cls, name):
        if name in ("__file__", "__path__"):
            return "/dev/null"
        elif name[0] == name[0].upper():
            mockType = Mock()
            mockType.__module__ = __name__
            return mockType
        else:
            return Mock()

    @classmethod
    def __getitem__(cls, key):
        return Mock()


MOCK_MODULES = [
    "appscript",
    "appscript.reference",
    "psutil",
    "xlplatform",
    "atexit",
    "aem",
    "osax",
]

if not sys.platform.startswith("win"):
    MOCK_MODULES += [
        "win32com",
        "win32com.client",
        "pywintypes",
        "pythoncom",
        "win32timezone",
        "win32com.server",
        "win32com.server.util",
        "win32com.server.dispatcher",
        "win32com.server.policy",
    ]

for mod_name in MOCK_MODULES:
    sys.modules[mod_name] = Mock()

# -- General configuration -----------------------------------------------------

# READTHEDOCS
html_baseurl = os.environ.get("READTHEDOCS_CANONICAL_URL", "")
html_context = {}
if os.environ.get("READTHEDOCS", "") == "True":
    html_context["READTHEDOCS"] = True

sys.path.insert(0, os.path.abspath("_ext"))

extensions = [
    "myst_parser",
    "sphinx.ext.autodoc",
    "sphinx.ext.napoleon",
    "sphinx.ext.mathjax",
    "sphinx.ext.extlinks",
    "sphinx_copybutton",
    "sphinx_design",
    "sphinx_llm.txt",
    "myst_docstrings",
]

templates_path = ["_templates"]
exclude_patterns = ["_build"]
master_doc = "index"

# -- Project information -------------------------------------------------------

project = "xlwings"
copyright = "Zoomer Analytics LLC"

import xlwings

version = xlwings.__version__
release = version

add_module_names = False
# Type hints and return types link as `Range`, not `xlwings.main.Range`
python_use_unqualified_type_names = True

# -- extlinks -----------------------------------------------------------------

extlinks = {"issue": ("https://github.com/xlwings/xlwings/issues/%s", "GH %s")}

# -- MyST configuration -------------------------------------------------------

myst_heading_anchors = 3
myst_enable_extensions = ["colon_fence", "linkify"]
myst_links_external_new_tab = True
myst_linkify_fuzzy_links = False

# -- Options for HTML output ---------------------------------------------------

html_theme = "furo"
html_static_path = ["_static"]
html_css_files = ["custom.css"]
html_show_sourcelink = False
# Link icon for the per-heading permalinks (Sphinx inserts this as raw HTML),
# replacing the default pilcrow. "currentColor" follows the light/dark theme.
html_permalinks_icon = (
    '<svg class="headerlink-icon" xmlns="http://www.w3.org/2000/svg"'
    ' viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2"'
    ' stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">'
    '<path d="M10 13a5 5 0 0 0 7.54.54l3-3a5 5 0 0 0-7.07-7.07l-1.72 1.71"/>'
    '<path d="M14 11a5 5 0 0 0-7.54-.54l-3 3a5 5 0 0 0 7.07 7.07l1.71-1.71"/>'
    "</svg>"
)
html_copy_source = False
html_title = "xlwings Documentation"
html_favicon = "_static/favicon.png"
html_extra_path = ["_static/opensource_licenses2.html"]
html_domain_indices = False
html_use_index = True
html_show_sphinx = False

html_theme_options = {
    "sidebar_hide_name": True,
    "top_of_page_buttons": [],
    "light_logo": "logo-light.svg",
    "dark_logo": "logo-dark.svg",
    "light_css_variables": {
        "color-brand-primary": "black",
        "color-brand-content": "#28a745",
        # Match visited links to unvisited ones; Furo's default is a purple
        # (#872ee0) that clashes with the green brand color.
        "color-brand-visited": "#28a745",
        "color-sidebar-caption-text": "#28a745",
        "sidebar-caption-font-size": "1em",
        "color-announcement-background": "#28a745",
    },
    "dark_css_variables": {
        "color-brand-primary": "white",
        # Lighter green than light mode (#28a745) for contrast on the dark
        # sidebar; drives the active sidebar item's accent bar/text/tint.
        "color-brand-content": "#3fbf5f",
        "color-brand-visited": "#3fbf5f",
        "color-announcement-background": "#28a745",
    },
    # "announcement": '<a href="https://lite.xlwings.org/" target="_blank"> xlwings Lite</a> is now available in the add-in store for free!</a>',
}

# -- LLM-friendly output -----------------------------------------------------
# Generates llms.txt, llms-full.txt, and a rendered Markdown version of each
# page alongside the regular HTML documentation.
llms_txt_description = (
    "Documentation for xlwings, a Python library to automate Excel, write "
    "user-defined functions, and build interactive Excel tools."
)
llms_txt_suffix_mode = "replace"

copybutton_prompt_text = r">>> |\.\.\. |\$ |In \[\d*\]: | {2,5}\.\.\.: | {5,8}: "
copybutton_prompt_is_regexp = True

suppress_warnings = ["misc.highlighting_failure"]

# -- Options for LaTeX output --------------------------------------------------

latex_elements = {
    "pointsize": "11pt",
    "printindex": "\\printindex",
}

latex_documents = [
    (
        "index_latex",
        "xlwings.tex",
        "xlwings - Make Excel Fly!",
        "Zoomer Analytics LLC",
        "manual",
        True,
    ),
]

latex_domain_indices = False
htmlhelp_basename = "xlwingsdoc"
texinfo_domain_indices = False

# Autodocs (recommended settings by Furo)

autodoc_typehints = "description"
autodoc_typehints_description_target = "documented_params"


def _hide_impl_signature(app, what, name, obj, options, signature, return_annotation):
    """Drop the ``impl`` parameter from class signatures.

    ``impl`` is the engine-specific implementation object and never meant to be
    passed by users, so it's noise in the docs. Classes that take nothing else
    render as ``class Foo`` without parentheses.
    """
    if what != "class":
        return None
    import inspect

    from sphinx.util.inspect import stringify_signature

    try:
        sig = inspect.signature(obj.__init__)
    except (TypeError, ValueError):
        return None
    params = list(sig.parameters.values())[1:]  # drop self
    if not any(p.name == "impl" for p in params):
        return None
    params = [p for p in params if p.name != "impl"]
    if not params:
        return "", return_annotation
    sig = sig.replace(parameters=params, return_annotation=inspect.Signature.empty)
    return stringify_signature(sig, show_annotation=False), return_annotation


def _widen_literals(annotation):
    """Replace `Literal["a", "b"]` with `str` (recursing into unions), so that
    the docs show the plain type; the docstrings list the accepted values."""
    import types
    import typing

    origin = typing.get_origin(annotation)
    if origin is typing.Literal:
        if all(isinstance(arg, str) for arg in typing.get_args(annotation)):
            return str
        return annotation
    if origin in (typing.Union, types.UnionType):
        args = tuple(_widen_literals(a) for a in typing.get_args(annotation))
        # Deduplicate while keeping order, e.g. Literal[...] | str -> str
        args = tuple(dict.fromkeys(args))
        return args[0] if len(args) == 1 else typing.Union[args]
    if origin is not None and typing.get_args(annotation):
        # Generics such as list[Literal[...] | str] -> list[str]
        args = tuple(_widen_literals(a) for a in typing.get_args(annotation))
        try:
            return origin[args]
        except TypeError:
            return annotation
    return annotation


def _widen_literal_typehints(
    app, what, name, obj, options, signature, return_annotation
):
    """Show `str` instead of `Literal[...]` in the parameter type fields.

    Runs after autodoc's own typehints recorder (same event, later priority)
    and rewrites the annotations it stored for this object.
    """
    import typing

    from sphinx.util.typing import stringify_annotation

    annotations = app.env.temp_data.get("annotations", {}).get(name)
    if not annotations:
        return None
    if what == "class":
        obj = getattr(obj, "__init__", None)
    try:
        hints = typing.get_type_hints(obj)
    except (AttributeError, NameError, TypeError, ValueError):
        return None
    for param, annotation in hints.items():
        widened = _widen_literals(annotation)
        if widened is not annotation and param in annotations:
            annotations[param] = stringify_annotation(widened)
    return None


def _add_setter_type(app, domain, objtype, contentnode):
    """Add a "Setter type" field to properties whose setter accepts a different
    type than the getter returns.

    Autodoc only shows the getter's return annotation as the property type,
    which hides e.g. the hex strings and ints that a color setter accepts.
    """
    if domain != "py" or objtype != "property":
        return
    import importlib
    import inspect
    import typing

    from docutils import nodes
    from sphinx import addnodes
    from sphinx.domains.python import _parse_annotation
    from sphinx.util.typing import stringify_annotation

    signature = contentnode.parent[0]
    module_name, fullname = signature.get("module"), signature.get("fullname")
    if not module_name or not fullname:
        return
    try:
        obj = importlib.import_module(module_name)
        for part in fullname.split("."):
            obj = getattr(obj, part)
        if not isinstance(obj, property) or obj.fset is None:
            return
        params = list(inspect.signature(obj.fset).parameters)
        setter_hints = typing.get_type_hints(obj.fset)
        getter_hints = typing.get_type_hints(obj.fget)
    except (AttributeError, NameError, TypeError, ValueError):
        return
    if len(params) < 2 or params[1] not in setter_hints:
        return
    setter_type = stringify_annotation(_widen_literals(setter_hints[params[1]]))
    if setter_type == stringify_annotation(getter_hints.get("return")):
        return
    field = nodes.field(
        "",
        nodes.field_name("", "Setter type"),
        nodes.field_body(
            "", nodes.paragraph("", "", *_parse_annotation(setter_type, app.env))
        ),
    )
    field_list = nodes.field_list("", field, classes=["simple"])
    # Keep the field ahead of any "Added in version" box so it stays with the
    # description
    for index, child in enumerate(contentnode):
        if isinstance(child, addnodes.versionmodified):
            contentnode.insert(index, field_list)
            return
    contentnode.append(field_list)


def _mark_type_fields(app, domain, objtype, contentnode):
    """Tag the "Return type" and "Setter type" field bodies so custom.css can
    render them in the signature's code font."""
    if domain != "py":
        return
    from docutils import nodes

    for field in contentnode.findall(nodes.field):
        if field[0].astext() in ("Return type", "Setter type"):
            field[1]["classes"].append("type-hint")


def _prepare_markdown_doctree(app, doctree, docname):
    """Fix internal references and asset URLs in generated Markdown."""
    if app.builder.name != "markdown":
        return

    from docutils import nodes

    for node in doctree.findall(nodes.reference):
        if node.get("refid") and not node.get("refuri"):
            node["internal"] = True

    for node in doctree.findall(nodes.image):
        image = app.env.images.get(node["uri"])
        if image:
            node["uri"] = f"/_images/{image[1]}"


def _add_markdown_twin_flag(app, pagename, templatename, context, doctree):
    """Flag pages that have a rendered Markdown twin.

    sphinx_llm emits a .md file per source document, but not for generated
    pages like genindex/search — linking those would 404. base.html uses this
    to decide whether to emit the text/markdown <link rel="alternate">.
    """
    context["has_markdown_twin"] = pagename in app.env.found_docs


def setup(app):
    app.connect("autodoc-process-signature", _hide_impl_signature)
    # Priority 600 runs after sphinx.ext.autodoc.typehints (default 500)
    app.connect("autodoc-process-signature", _widen_literal_typehints, priority=600)
    app.connect("object-description-transform", _add_setter_type)
    app.connect("object-description-transform", _mark_type_fields, priority=600)
    app.connect("doctree-resolved", _prepare_markdown_doctree)
    app.connect("html-page-context", _add_markdown_twin_flag)
