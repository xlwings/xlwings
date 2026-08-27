# Office.js ("remote") engine — NotImplemented API surface

Status as of 2026-08-27, based on comparing `xlwings/base_classes.py` against
`xlwings/pro/_xlremote.py`. Check off items as they get implemented (or mark
them `n/a` if they can't/shouldn't be supported in Office.js).

Note on write-only properties: the setter works (via JSON actions), but the
getter raises `NotImplementedError` because the value isn't included in the
payload sent to Python.

## Missing classes (no implementation at all)

- [ ] `Chart`
  - [ ] `parent`
  - [ ] `name`
  - [ ] `chart_type` (get/set)
  - [ ] `left` / `top` / `width` / `height` (get/set)
  - [ ] `set_source_data()`
  - [ ] `delete()`
  - [ ] `to_png()`
  - [ ] `to_pdf()`
- [ ] `Charts`
  - [ ] `add()`
- [ ] `Shape`
  - [ ] `parent`
  - [ ] `name` (get/set)
  - [ ] `type`
  - [ ] `left` / `top` / `width` / `height` (get/set)
  - [ ] `index`
  - [ ] `text`
  - [ ] `font`
  - [ ] `characters`
  - [ ] `activate()`
  - [ ] `delete()`
  - [ ] `scale_width()`
  - [ ] `scale_height()`
- [ ] `Characters`
  - [ ] `text`
  - [ ] `font`
- [ ] `Note`
  - [ ] `text` (get/set)
  - [ ] `delete()`
- [ ] `PageSetup`
  - [ ] `print_area` (get/set)

## App

- [ ] `calculation` (get/set)
- [ ] `calculate()`
- [ ] `cut_copy_mode` (get/set)
- [x] `display_alerts` (get/set) — stored but unused; Office.js shows no alerts
- [ ] `enable_events` (get/set)
- [ ] `interactive` (get/set)
- [ ] `screen_updating` (get/set)
- [ ] `status_bar` (get/set)
- [ ] `path`
- [ ] `startup_path`
- [ ] `version`
- [ ] `quit()`

## Book

- [ ] `save()`
- [ ] `to_pdf()`

## Sheet

- [x] `used_range` — the client sends the used range's address as
      `used_range_address` (an exception to the "no new payload data" rule:
      the `values` payload is anchored at A1, so the real top-left corner
      can't be recovered from it). Clients that don't send it fall back to
      deriving the extent from the shape of `values`, which then starts at A1
      and raises on an unloaded lazy (async) book.
- [ ] `visible` (get/set)
- [ ] `charts`
- [ ] `shapes`
- [ ] `page_setup`
- [ ] `autofit()`
- [ ] `copy()`
- [ ] `select()`
- [ ] `to_html()`

## Range

Decision (2026-08-20): only implement **setters and action methods** for now.
Getters are deferred, since they would require adding more data to the JSON
payload that the Office.js client sends with every request.

Done — setters:

- [x] `formula` (setter) — `setFormula`
- [x] `formula2` (setter) — delegates to `formula`; Office.js `range.formulas`
      already writes dynamic array formulas
- [x] `formula_array` (setter) — writes the formula to the top-left cell only so
      it spills; Office.js can't write legacy CSE arrays (`savedAsArray` is
      read-only)
- [x] `column_width` (setter) — `setColumnWidth`; the value is passed through
      as points, which is what Office.js uses. Note that this differs from the
      COM API, which measures column widths in characters
- [x] `row_height` (setter) — `setRowHeight`; points on both sides
- [x] `wrap_text` (setter) — `setWrapText`

Done — action methods:

- [x] `merge()` — `rangeMerge` (needed `App.display_alerts`, see below)
- [x] `unmerge()` — `rangeUnmerge`
- [x] `autofill()` — `rangeAutofill`; rejects a destination on another sheet,
      since Office.js resolves it against the source's sheet

Not implementable — `paste()`:

- `paste()` reads the system clipboard, which Office.js has no API for.
  `range.copyFrom()` needs an explicit source range and is already exposed as
  `Range.copy()`, so there's nothing to map `paste()` onto. Left raising.

Also implemented along the way:

- [x] `App.display_alerts` — required by `Range.merge()`, which wraps itself in
      `app.properties(display_alerts=False)`. Stored but unused, as Office.js
      never shows those alerts.

Done — async getters (fetched on demand, so the sync payload stays small):

- [x] `get_formula()` — `js.xlwings.getRangeFormulas`. Returns a scalar for a
      single cell and a nested list for a range, like the COM API. Works on a
      lazy (async) book without loading values, since it doesn't read the
      payload. The sync `formula` / `formula2` getters still raise, but now
      point at `get_formula()`. Lite-only, like `get_value()`.

Deferred — getters (would require more data in the JSON payload):

- `formula_array` (getter)
- `column_width` / `row_height` / `wrap_text` (getters)
- `color` / `number_format` (getters; setters already work)
- `left` / `top` / `width` / `height`
- `current_region`
- `merge_area` / `merge_cells`
- `hyperlink` (getter; `add_hyperlink()` already works)
- `note`
- `table`
- `characters`
- `rows` / `columns`
- `get_address()`

Deferred — methods that need to return data or aren't pure JSON actions:

- `copy_picture()`
- `to_pdf()`
- `paste()` (see above — no Office.js clipboard API)

## Font

All setters work; implement the getters:

- [ ] `bold` (getter)
- [ ] `italic` (getter)
- [ ] `size` (getter)
- [ ] `color` (getter)
- [ ] `name` (getter)

## Picture

- [ ] `left` (get/set)
- [ ] `top` (get/set)
- [ ] `lock_aspect_ratio` (get/set)
- [ ] `index()`

## Table

- [ ] `display_name` (get/set)
- [ ] `insert_row_range`
- [ ] `show_table_style_first_column` (get/set)
- [ ] `show_table_style_last_column` (get/set)
- [ ] `show_table_style_row_stripes` (get/set)
- [ ] `show_table_style_column_stripes` (get/set)

## Name

- [ ] `name` (setter; getter works)
- [ ] `refers_to` (setter; getter works)

## Lite-only (work in xlwings Lite, raise on remote/server)

These are implemented but gated on `sys.platform == "emscripten"`; decide
whether the remote engine should support them too:

- [ ] `App.get_selection()`
- [ ] `Books.get_active()`
- [ ] `Sheets.get_active()`
- [ ] `Book.flush()`
- [ ] `Book.load()`
- [ ] `Sheet.load()`
- [ ] `Range.get_value()`
- [ ] `Range.get_formula()`
