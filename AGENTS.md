# Repository Instructions

## Documentation

- Write all documentation and Python docstrings using MyST Markdown.
- Never use reStructuredText syntax in documentation prose or Python docstrings, including double-backtick inline literals, roles such as `:attr:` and `:meth:`, or directives such as `.. note::` and `.. py:method::`.
- Use single backticks for inline code, MyST roles such as `{attr}` and `{meth}`, and fenced MyST directives such as the `note` directive instead.
- API reference stub pages are the one exception: follow the established neighboring pages and place `.. autoclass::` or `.. automodule::` with `:members:` inside a fenced `{eval-rst}` directive. This compatibility wrapper produces the standard Sphinx Python class and member rendering used throughout the API reference.
- Do not replace the established `{eval-rst}` API wrapper with direct fenced `{autoclass}`, `{automodule}`, or `{automethod}` directives. In this repository's current Sphinx pipeline, those directives can emit literal reStructuredText such as `.. py:method::` into generated Markdown and HTML.
- Before adding or changing an API reference page, inspect a neighboring page for the same kind of object and match its source structure and rendered output.
- Never hard-wrap prose. Write each paragraph, list item, and table cell as a single long line and let the editor soft-wrap.
- After adding or changing API documentation, build the docs and inspect the generated Markdown and HTML for the changed page. A successful Sphinx exit is insufficient: fail the review if generated output contains literal `.. py:`, `.. eval-myst::`, or other leaked parser directives.
- Describe ordinary synchronized reads plainly. Do not call returned objects "live," "on-demand," or "point-in-time snapshots" unless those terms express a user-visible semantic distinction; fetching current state is analogous to a sync operation and does not require special vocabulary.
