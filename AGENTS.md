# Repository Instructions

## Documentation

- Write all documentation and Python docstrings using MyST Markdown.
- Never use reStructuredText syntax, including roles such as `:attr:` and
  `:meth:`, directives such as `.. note::`.
- Use MyST roles such as `{attr}` and `{meth}`, and MyST directives such as
  ```` ```{note} ```` instead.
- Describe ordinary synchronized reads plainly. Do not call returned objects
  "live," "on-demand," or "point-in-time snapshots" unless those terms express
  a user-visible semantic distinction; fetching current state is analogous to a
  sync operation and does not require special vocabulary.
