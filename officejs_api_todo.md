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
- [x] `Shape` — the client now sends a per-sheet `shapes` array (one load
      covers pictures too, since a picture is a shape whose type is `Image`),
      so the getters are payload reads and the setters queue `setShape*` JSON
      actions
  - [x] `parent`
  - [x] `name` (get/set)
  - [x] `type` — Office.js' `ShapeType` is only
        `Unsupported`/`Image`/`GeometricShape`/`Group`/`Line`, far coarser than
        the desktop engines' ~30 types, so it's passed through as-is rather
        than mapped onto names implying a precision Office.js doesn't have
  - [x] `left` / `top` / `width` / `height` (get/set)
  - [x] `index` — impl-only, as on `Table` and `Picture`
  - [x] `text` (setter) — the getter raises: the payload carries geometry, not
        text
  - [ ] `font` — needs `Characters`/`TextRange` plumbing; raises for now
  - [ ] `characters` — needs the `Characters` class
  - [x] `activate()` — **n/a**: Office.js has no way to activate or select a
        shape. `setZOrder(bringToFront)` is a *different* operation, so it
        raises rather than doing something else under the same name
  - [x] `delete()`
  - [x] `scale_width()` / `scale_height()` — `scaleShape`, mapping
        `relative_to_original_size` onto `ShapeScaleType`
        (`CurrentSize`/`OriginalSize`) and xlwings' `scale_from_*` onto
        `ShapeScaleFrom`; an unknown anchor raises `ValueError`
- [ ] `Characters`
  - [ ] `text`
  - [ ] `font`
- [ ] `Note`
  - [ ] `text` (get/set)
  - [ ] `delete()`
- [x] `PageSetup`
  - [x] `print_area` (get/set) — the client sends the sheet's print area as
        `print_area`, so the getter is a payload read and the setter queues
        `setPrintArea`. Office.js reports it as a `RangeAreas` (one or more
        rectangles), which the client joins with commas the way Excel writes
        them, and `None` when there is none --- matching win32, which maps its
        empty string to `None`. Clearing (`print_area = None`, which the public
        API documents) passes an empty string to `setPrintArea`, since
        Office.js has no explicit clear method. Sent in lazy mode too, like
        `used_range_address`, so it works on an async book

## App

Verified against the `Excel.Application` typings, whose entire surface is:
`calculationMode`, `calculationState`, `calculationEngineVersion`,
`cultureInfo`, `decimalSeparator`, `thousandsSeparator`, `useSystemSeparators`,
`iterativeCalculation`, `activeWindow`, `windows`, plus the methods
`calculate()`, `checkSpelling()`, `enterEditingMode()`, `union()`,
`suspendApiCalculationUntilNextSync()` and
`suspendScreenUpdatingUntilNextSync()`. Everything not on that list is `n/a`.

- [x] `calculation` (get/set) — maps to `calculationMode`, with
      `semiautomatic` ↔ `AutomaticExceptTables`. The client now sends it as
      `book.calculation`, loaded in an existing `context.sync()` so it costs no
      extra round-trip; the setter queues `setCalculation`. A payload without
      the field (an older client) raises rather than `KeyError`
- [x] `calculate()` — `calculate` JSON action, calling
      `application.calculate(Excel.CalculationType.full)`. No payload cost
- [x] `cut_copy_mode` (get/set) — **n/a**: Office.js has no clipboard API
- [x] `display_alerts` (get/set) — stored but unused; Office.js shows no alerts
- [x] `enable_events` (get/set) — **n/a**: no equivalent property
- [x] `interactive` (get/set) — **n/a**: no equivalent property
- [x] `screen_updating` (set only) — Office.js only has
      `suspendScreenUpdatingUntilNextSync()`, so setting `False` queues that
      call and setting `True` is a no-op: the suspension ends at the next sync
      by itself. The getter raises, as there's no flag to read back. The
      public docstring documents the narrower Server/Lite behaviour
- [x] `status_bar` (get/set) — **n/a**: no equivalent property
- [x] `path` — **n/a**: an add-in has no access to the installation's paths
- [x] `startup_path` — **n/a**: as `path`
- [x] `version` — **n/a**: only `calculationEngineVersion` exists, which is not
      the application version
- [x] `quit()` — **n/a**: an add-in can't close the Excel application

The `n/a` members raise `NotImplementedError` with the reason. `path`,
`startup_path` and `version` are read-only in the public API, so they define no
setter --- assigning to them raises `AttributeError` there, as on every other
backend.

## Book

Verified against the `Excel.Workbook` typings: its methods are `save()`,
`close()`, `focus()`, the `getActive*` family, `getSelectedRange(s)` and
`insertWorksheetsFromBase64()`. There is no PDF export anywhere on the type,
which settles `Sheet.to_pdf()` and `Range.to_pdf()` below as well.

- [x] `save()` — no-argument form only, mapping to
      `workbook.save(Excel.SaveBehavior.save)`. `path` raises (Office.js has no
      SaveAs, so silently saving in place would be wrong) and so does
      `password` (no API for it). Like every JSON action the save lands after
      the script returns, which the public docstring now notes
- [x] `to_pdf()` — **n/a**: Office.js has no PDF export API

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
- [ ] `charts` — `Excel.Worksheet` has `charts`, so this is reachable, but it
      returns a collection whose element class doesn't exist yet. Belongs with
      the `Chart` work above rather than here
- [x] `shapes` — returns the `Shapes` collection above
- [x] `page_setup` — returns the `PageSetup` above
- [x] `autofit()` — `setSheetAutofit` JSON action. Separate from the range-level
      `setAutofit`, which resolves its target through `getRange()` using the
      action's row/column coordinates; a sheet-level action has none, so it
      autofits `sheet.getRange()` (the whole sheet) instead
- [x] `copy()` — `copySheet` JSON action, mapping onto
      `Worksheet.copy(positionType, relativeTo)`. The public `Sheet.copy()`
      identifies the new sheet by diffing sheet names before and after the
      call, so the impl inserts the copy into the local api list synchronously
      (as `Sheets.add()` does) rather than relying on the queued action; it
      predicts Excel's `"<name> (n)"` naming and the client renames the copy to
      match, so both sides agree. Copying to a *different* book raises:
      Office.js positions the copy within the same workbook only
- [x] `select()` — Office.js has no separate select, so this activates the
      sheet, reusing the existing `activateSheet` action
- [x] `to_html()` — **n/a**: Office.js has no HTML export API

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

**Settled — where a new mode lands.** `getRangeData` is defined only in
xlwings-server and reaches Lite through `globalThis.xlwings`, which Lite reads
rather than defines. So a new mode is a **one-repo change** (plus the Python
wrapper); Lite picks it up when it bumps its xlwings-server dependency. The
`range_read_unavailable` guard in `static/js/wingman/workbook.js` is defensive
coding, not a second implementation.

`mode` now takes a **list of keys** — `getRangeData(sheet, address,
["number_format", "color"])` — so several properties come back in one
round-trip. The old single strings are gone. Lite's Wingman `read_range` tool
still shows the model a `"values"`/`"formulas"`/`"both"` enum, but that's its
own vocabulary: it translates to read keys at the call site
(`RANGE_MODE_READ_KEYS` in `static/js/wingman/workbook.js`) rather than the
shared plumbing knowing what `"both"` means.

Done — Group A getters (plain Office.js `range.*` properties):

- [x] `formula_array` (getter) — `get_formula_array()`. A single string or
      `None`, like `formula_array` on the other engines --- *not* a per-cell
      matrix, so unlike `get_formula()` it doesn't go through
      `AdjustDimensionsStage`
- [x] `column_width` / `row_height` / `wrap_text` (getters) — `column_width`
      converts Office.js points back to xlwings' characters. The setter can
      measure the workbook's real digit width because it resets the column
      first; a getter mustn't mutate the sheet, so it assumes 7px (Calibri 11)
      and is off by a couple of percent for other Normal-style fonts. Office.js
      returns `null` for a non-uniform range, which passes through as `None`
- [x] `color` / `number_format` (getters) — `color` converts to the RGB tuple
      the other backends return, `None` when unset. Office.js may report a
      *named* HTML colour ("orange") rather than `#RRGGBB`, which the client
      normalizes first, since `hex_to_rgb()` would raise on it.
      `number_format` is a single string (or `None` when the cells disagree),
      matching COM's scalar `NumberFormat`; Office.js reports a per-cell
      matrix, so the client collapses it
- [x] `left` / `top` / `width` / `height` — these had no sync property on this
      engine at all; they now raise pointing at the async version like the rest

Done — Group B getters (resolved through Office.js *method* calls, returning
addresses that the `Range`/`Table` object is then built from locally):

- [x] `current_region` — `get_current_region()`, via `getSurroundingRegion()`
- [x] `merge_area` / `merge_cells` — both from
      `getMergedAreasOrNullObject()`. Office.js reports every merged area
      overlapping the range, and none at all for an unmerged cell, whereas COM
      echoes the cell back --- so `get_merge_area()` falls back to the range
      itself, matching `Range.MergeArea`. `get_merge_cells()` is tri-state like
      COM's `Range.MergeCells`: `True` when the whole range is merged, `False`
      when none of it is, and `None` when it's only partly merged (compared by
      cell count, since Office.js has no such flag)
- [x] `table` — `get_table()`, via `getTables(false)`. Returns `None` when the
      range isn't in a table. The client reports the table's *name*, which the
      Python side resolves to its position, since `Table`'s constructor indexes
      the sheet's tables list
- [x] `hyperlink` (getter) — `get_hyperlink()`, from `range.hyperlink`
      (`RangeHyperlink.address`, falling back to `documentReference` for
      in-workbook targets). It's a *public* async method rather than only an
      impl one, because `Range.hyperlink` in `main.py` reads `self.formula`
      first to handle `HYPERLINK()` formulas --- the sync property this engine
      raises on. `get_hyperlink()` reproduces both branches, and raises
      "The cell doesn't seem to contain a hyperlink!" when there is none, as
      both desktop engines do rather than returning `None`
- [x] `get_address()` — **no fetch needed**: the engine already knows the
      range's coordinates (the `address` property is built from them), so this
      is local string formatting. Honours `row_absolute` / `col_absolute` and
      the `external` prefix, quoting it when the book or sheet name contains a
      space, as Excel does
- [x] `rows` / `columns` — **already worked**: `main.py` builds `RangeRows` /
      `RangeColumns` from the range itself and never calls the impl. They were
      listed here by mistake

Still deferred — getters that need more than a mode:

- `note` — needs the `Note` class
- `characters` — needs the `Characters` class

Two notes on shape and scope:

- ~~Anything returning a cell matrix (`formula_array`, and the `color` /
  `number_format` getters) should run through `AdjustDimensionsStage`~~ ---
  **wrong**: on every engine `formula_array`, `color` and `number_format` are
  *scalars*, not per-cell matrices. Only `get_formula()` is a matrix and goes
  through that stage. The non-uniform case is expressed as `None`, which both
  COM and Office.js do.
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

- [x] `bold` (getter)
- [x] `italic` (getter)
- [x] `size` (getter)
- [x] `color` (getter)
- [x] `name` (getter)

      One `font` read key covers all five: they come from a single
      `range.format.font` object, so `_get_font()` fetches them together and
      each `get_*()` picks its attribute --- there's nothing to gain from
      fetching them one at a time. `get_color()` returns the RGB tuple the
      other backends return (and normalizes a named HTML colour first, like
      the range fill). A range whose cells disagree reports `None`, which is
      what the public annotations already said.

`Font.append_json_action()` still raises for any parent that isn't a `Range`,
so `Shape.font` and `Characters.font` need that path before their setters work
--- but it now raises with the reason rather than a bare `NotImplementedError`,
and the getters raise the same way. That's unblocked by the `Shape` and
`Characters` classes above, not by anything in this section.

## Picture

- [x] `left` (get/set)
- [x] `top` (get/set)
- [x] `lock_aspect_ratio` (get/set)

      Case 1: the three properties are now loaded into the per-sheet
      `pictures` payload (`Excel.Shape` has `left`, `top` and
      `lockAspectRatio`), so the getters are payload reads and the setters
      queue `setPictureLeft` / `setPictureTop` /
      `setPictureLockAspectRatio`. `Pictures.add()` seeds them, as
      `Sheets.add()` and `Tables.add()` had to.

      Fixed alongside: the existing `width` / `height` setters queued their
      action without updating the local api dict, so reading them back in the
      same script returned the old value.

`index` is already implemented (as a property, not a method), along with
`name`, `width`, `height`, `delete()` and `update()`.

## Table

Mostly case 1: the payload already builds a `tables` array per sheet, and the
neighbouring `show_headers` / `show_totals` / `show_autofilter` flags are read
straight off it. These need the same fields added on the client side, then the
getters are one-liners and the setters are JSON actions:

- [x] `display_name` (get/set) --- aliases `name`. Office.js' `Excel.Table` has
      no `displayName` property (its full property set is `name`, `style`,
      `showHeaders`, `showTotals`, `showFilterButton`, `highlightFirstColumn`,
      `highlightLastColumn`, `showBandedRows`, `showBandedColumns`), and the
      two are equivalent in practice: on macOS setting `display_name` changes
      the name too, and Office Scripts dropped the distinction as well.
      Aliasing keeps scripts that use `display_name` portable across backends
      instead of raising for a distinction this host doesn't make.
- [x] `show_table_style_first_column` (get/set)
- [x] `show_table_style_last_column` (get/set)
- [x] `show_table_style_row_stripes` (get/set)
- [x] `show_table_style_column_stripes` (get/set)

      The four flags are named differently in Office.js: `highlightFirstColumn`,
      `highlightLastColumn`, `showBandedRows` and `showBandedColumns`. They're
      now loaded into the per-sheet `tables` payload under their xlwings names,
      so the getters are payload reads; the setters queue
      `showTableStyle*` JSON actions. `Tables.add()` seeds them (plus
      `show_autofilter`, which it was missing) with Excel's defaults.
- [x] `insert_row_range` --- **n/a**: Office.js' `Excel.Table` has no
      `InsertRowRange` equivalent; its only range accessors are `getRange()`,
      `getDataBodyRange()`, `getHeaderRowRange()` and `getTotalRowRange()`,
      all of which are already implemented. Raises rather than returning
      `None`, which is the documented answer for a non-empty table and so
      would be indistinguishable from a real result.

## Name

- [x] `name` (setter) — **n/a**: `Excel.NamedItem.name` is read-only in
      Office.js, so a named item can't be renamed. Delete-and-recreate would
      change its identity and drop its comment and visibility, so the setter
      raises rather than doing that implicitly. The getter already worked
- [x] `refers_to` (setter) — `setNameRefersTo` JSON action writing
      `NamedItem.formula`, the writable counterpart of the read-only `name`.
      Handles book- and sheet-scoped names (passing `scope_sheet_index` the way
      `nameDelete` does). The getter is computed from `sheet_index` and
      `address`, so the setter updates those rather than storing the string;
      an unknown sheet raises `ValueError`

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
