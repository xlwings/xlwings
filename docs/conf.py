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
html_show_sourcelink = False
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
        "color-sidebar-caption-text": "#28a745",
        "sidebar-caption-font-size": "1em",
        "color-announcement-background": "#28a745",
    },
    "dark_css_variables": {
        "color-brand-primary": "white",
        "color-announcement-background": "#28a745",
    },
    "announcement": '<a href="https://lite.xlwings.org/" target="_blank"> xlwings Lite</a> is now available in the add-in store for free!</a>',
}

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
autodoc_typehints_description_target = "documented"