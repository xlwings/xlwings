# Repository Instructions

## Documentation

- Write all documentation and Python docstrings using MyST Markdown.
- Never use reStructuredText syntax, including double-backtick inline literals, roles such as `:attr:` and `:meth:`, or directives such as `.. note::`, `.. autoclass::`, and `.. py:method::`.
- Use single backticks for inline code, MyST roles such as `{attr}` and `{meth}`, and fenced MyST directives such as the `note` directive instead.
- Do not use fenced `{autoclass}`, `{automodule}`, or `{automethod}` directives on new Markdown pages. In this repository's current Sphinx pipeline, they can emit literal reStructuredText such as `.. py:method::` into generated Markdown and HTML. Manually document new API classes and members with Markdown headings and signatures until the pipeline has a verified MyST-native autodoc path. Existing `eval-rst` API pages are legacy compatibility files and are not examples to copy.
- Never hard-wrap prose. Write each paragraph, list item, and table cell as a single long line and let the editor soft-wrap.
- After adding or changing API documentation, build the docs and inspect the generated Markdown and HTML for the changed page. A successful Sphinx exit is insufficient: fail the review if generated output contains literal `.. py:`, `.. eval-myst::`, or other leaked parser directives.
- Describe ordinary synchronized reads plainly. Do not call returned objects "live," "on-demand," or "point-in-time snapshots" unless those terms express a user-visible semantic distinction; fetching current state is analogous to a sync operation and does not require special vocabulary.
