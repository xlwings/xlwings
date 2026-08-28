# Office.js ("remote") engine — NotImplemented API surface

Status as of 2026-08-28, based on comparing `xlwings/base_classes.py` against
`xlwings/pro/_xlremote.py`. Check off items as they get implemented (or mark
them `n/a` if they can't/shouldn't be supported in Office.js).

## The general pattern: setters are sync, getters are async

This applies to **every** class here, not just `Range`, so it's worth stating
once up front.

Writing is already solved everywhere: a setter calls `append_json_action()`,
which queues a JSON action the client applies on its next turn. Nothing has to
come back, so setters stay synchronous and work on both Lite and remote. That's
why most setters below are cheap to add.

Reading is the hard direction, because the data has to travel back to Python.
There are two sources, and which one applies decides the work:

1. **Already in the payload** — the client sends `values`, plus per-sheet
   metadata (`name`, `visibility`, `used_range_address`) and the `tables` and
   `pictures` arrays with a fixed set of properties. Getters over this data are
   synchronous and easy: `Table.name`, `Picture.width`, `Sheet.name` all just
   read `self.api[...]`. Some unimplemented getters fall here too — e.g.
   `Sheet.visible` (the payload carries `visibility`) and the `Table`
   `show_table_style_*` flags (a few extra fields on the existing tables array).
2. **Not in the payload** — everything else. The fix is *not* to grow the
   payload, since that cost is paid on every request whether or not anything
   reads the property. Instead add an **async `get_*()` method** that fetches on
   demand, the way `Range.get_value()` and `Range.get_formula()` do, and leave
   the sync property raising with a message pointing at the async version.

So the rule of thumb for anything below: **a getter is either a payload field
away or an async method away.** `Chart` and `Shape` aren't in the payload at
all, so their getters are all case 2 — while `App` is mostly neither, since
`version`, `path` and friends aren't range data and would need their own
client-side call (or to be dropped as `n/a`, like the process-level members).

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
      can't be recovered from it). It's metadata rather than cell values, so
      the client sends it in lazy mode too and `used_range` works on an async
      book without loading any values. Clients that don't send it fall back to
      deriving the extent from the shape of `values`, which then starts at A1
      and raises on an unloaded lazy (async) book.
- [x] `visible` (get/set) — case 1: the getter reads the payload's
      `visibility` field; the setter queues a `setSheetVisibility` JSON action.
      Office.js' third state, `VeryHidden`, maps to `False` like `Hidden` does,
      since xlwings' public API is a bool. `Sheets.add()` seeds `visibility`
      on the new sheet's api dict, so the getter works before the next
      round-trip.
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
- [x] `formula_array` (setter) — `setFormulaArray`; uses Office.js'
      desktop-only `formulaArray` API and raises on unsupported hosts
- [x] `column_width` (setter) — `setColumnWidth`; accepts xlwings' public
      character units and converts them to Office.js points using the sheet's
      standard character and point widths
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

- [x] `get_formula()` — `js.xlwings.getRangeData(..., "formulas")`. The result
      goes through `AdjustDimensionsStage`, the same stage `.value` reads use,
      so the shape rules match: scalar for a single cell, flat list for a
      1-by-n or n-by-1 range, nested list otherwise --- and `options(ndim=...)`
      works too. Works
      on a lazy (async) book without loading values, since it doesn't read the
      payload. The sync `formula` / `formula2` getters still raise, but now
      point at `get_formula()`. Lite-only, like `get_value()`.

Deferred — getters. These are case 2 from the header: implement as async
`get_*()` methods, keeping the sync property raising with a message that points
at the async version (the `formula` / `get_formula()` pattern).

For ranges specifically the client side is already generic:
`js.xlwings.getRangeData(sheetName, address, mode)` loads whichever Office.js
properties `mode` selects and returns them alongside the range's address and
row/column counts. Today `mode` accepts `"values"`, `"formulas"` and `"both"`
(see `rangeReadProperties()` in xlwings-server's `custom-scripts/index.js`), so
most of the list below is a new `mode` plus a thin async wrapper on the Python
side — not new plumbing:

**Open question — where a new `mode` has to land.** `getRangeData` is defined in
xlwings-server. The Lite addin doesn't define it: it reads
`globalThis.xlwings?.getRangeData` and raises `range_read_unavailable` if it's
missing (`static/js/wingman/workbook.js`). So adding a mode may be one change or
two coordinated ones across repos, depending on how Lite is provisioned with
that function. Worth settling before the first async getter, since every one of
them goes through this path. TODO: confirm and write down the mechanism.

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

Two notes on shape and scope:

- Anything returning a cell matrix (`formula_array`, and the `color` /
  `number_format` getters) should run through `AdjustDimensionsStage` the way
  `get_formula()` does, so the result matches how `.value` reads shape and
  `options(ndim=...)` keeps working.
- `current_region`, `merge_area`, `rows` / `columns` and `table` return
  *ranges*, not data. The async fetch only needs to resolve the address; the
  `Range` object itself is then built synchronously from it.

Deferred — methods that need to return data or aren't pure JSON actions:

- `copy_picture()`
- `to_pdf()`
- `paste()` (see above — no Office.js clipboard API)

## Font

All setters work. The getters have the same constraint as the `Range` getters
above — font attributes aren't in the payload — so they'd follow the same async
route (a `getRangeData` mode that loads `range.format.font`):

- [ ] `bold` (getter)
- [ ] `italic` (getter)
- [ ] `size` (getter)
- [ ] `color` (getter)
- [ ] `name` (getter)

`Font.append_json_action()` also raises for any parent that isn't a `Range`
(see its `TODO: support Shape and getters`), so `Shape.font` and
`Characters.font` need that path before their setters can work either.

## Picture

- [ ] `left` (get/set)
- [ ] `top` (get/set)
- [ ] `lock_aspect_ratio` (get/set)

`index` is already implemented (as a property, not a method), along with
`name`, `width`, `height`, `delete()` and `update()`.

## Table

Mostly case 1: the payload already builds a `tables` array per sheet, and the
neighbouring `show_headers` / `show_totals` / `show_autofilter` flags are read
straight off it. These need the same fields added on the client side, then the
getters are one-liners and the setters are JSON actions:

- [ ] `display_name` (get/set)
- [ ] `show_table_style_first_column` (get/set)
- [ ] `show_table_style_last_column` (get/set)
- [ ] `show_table_style_row_stripes` (get/set)
- [ ] `show_table_style_column_stripes` (get/set)

- [ ] `insert_row_range` — returns a `Range`, so it needs an address rather
      than a flag

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

Note that every async getter added per the `Range` section above lands in this
list too: the `sys.platform != "emscripten"` guard makes it Lite-only by
default. So the more of the deferred getters get implemented this way, the
wider the Lite/remote gap grows — worth deciding the remote story once, rather
than per getter.
