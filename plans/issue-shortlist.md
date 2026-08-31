# Issue shortlist

Shortlist drawn from the 398 open issues on 2026-08-31, ranked by reactions,
comments, and recency.

## Recurring bugs — several issues each, so fixing one closes many

### 1. `xlwings vba edit` crashes on non-UTF-8 VBA modules

[#2654](https://github.com/xlwings/xlwings/issues/2654),
[#2692](https://github.com/xlwings/xlwings/issues/2692),
[#2683](https://github.com/xlwings/xlwings/issues/2683),
[#2588](https://github.com/xlwings/xlwings/issues/2588),
[#2148](https://github.com/xlwings/xlwings/issues/2148)

Five separate reports of `UnicodeDecodeError` when modules contain accented
characters (French, German, Portuguese). VBA exports modules in the legacy
system codepage, not UTF-8, so reading with a fallback to `cp1252`/locale
encoding (or `charset-normalizer`) would likely fix all of them. Best
value-for-effort on the whole list.

### 2. Type-hint analysis regression in UDFs since v0.32.0

[#2666](https://github.com/xlwings/xlwings/issues/2666) (labeled bug) plus
[#2586](https://github.com/xlwings/xlwings/issues/2586) (Union type hints)

A regression that breaks previously working subs deserves priority over any
feature, and both point at the same type-hint parsing code.

### 3. OneDrive/SharePoint path resolution

[#2160](https://github.com/xlwings/xlwings/issues/2160) (14 comments),
[#2652](https://github.com/xlwings/xlwings/issues/2652),
[#2576](https://github.com/xlwings/xlwings/issues/2576),
[#2365](https://github.com/xlwings/xlwings/issues/2365)

The `fullname`-is-a-URL problem is probably the most common real-world
friction for corporate users. A hardened resolver (registry-based mount-point
lookup on Windows, config fallback) would retire a whole class of support
threads.

## High-signal individual bugs

### 4. UDFs disappear when a second workbook is opened

[#2608](https://github.com/xlwings/xlwings/issues/2608)

17 comments, the most-discussed issue of the last two years. Worth a real
investigation even if the fix is just documented behavior.

### 5. `view()` / `Range.value` fails on pandas `Period` dtype

[#2741](https://github.com/xlwings/xlwings/issues/2741)

Newest issue; likely a small converter fix (stringify Periods) in a
well-understood code path.

### 6. Single quotes in sheet names break defined names

[#2649](https://github.com/xlwings/xlwings/issues/2649)

Classic quoting/escaping bug, small and self-contained.

## Features with sustained demand

### 7. Cell formatting API

[#559](https://github.com/xlwings/xlwings/issues/559) (7 reactions, open
since 2016) and [#2324](https://github.com/xlwings/xlwings/issues/2324)
(19 comments, `number_format` on more objects)

The most-requested feature overall, and it dovetails with the officejs-api
branch — formatting primitives (column widths, shapes) already exist for that
engine, so a cross-engine formatting surface would land on prepared ground.

### 8. Lazy imports

[#2651](https://github.com/xlwings/xlwings/issues/2651)

`import xlwings` currently pulls in numpy/pandas/polars when installed. Cheap
to fix with deferred converter registration, and startup time matters most
exactly in wasm/Office.js contexts.

### 9. `book.refresh_all()`

[#1592](https://github.com/xlwings/xlwings/issues/1592)

Tiny API addition, repeatedly asked for, trivially maps to `RefreshAll` on
both COM engines.

## Deliberately skipped for now

- PivotTables ([#191](https://github.com/xlwings/xlwings/issues/191)) has
  decade-long demand but is a large API-design project, not a shortlist item.
- The zombie-process cluster
  ([#1789](https://github.com/xlwings/xlwings/issues/1789),
  [#277](https://github.com/xlwings/xlwings/issues/277)) is largely
  pywin32-lifecycle territory with workarounds documented.
- A good third of the open issues are usage questions or environment problems
  (antivirus, Windows Defender, PyInstaller) that are better handled by docs
  than code.

## Suggested starting point

The VBA-edit encoding fix (#1), the type-hint regression (#2), and the Period
converter (#5) — all small, all bugs, and together they close eight or nine
open issues.
