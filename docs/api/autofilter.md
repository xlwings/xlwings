# AutoFilter

Use {attr}`Range.autofilter <xlwings.Range.autofilter>` or {attr}`Table.autofilter <xlwings.Table.autofilter>` to apply and clear filters without changing cell values, formulas, formats, or table identity. Fields are one-based positions relative to the range or table.

```python
data = sheet["A1:C100"]

# Keep rows whose first field is East or West
data.autofilter.apply_values(1, ["East", "West"])

# Keep rows whose third field is between 10 and 20, inclusive
data.autofilter.apply_comparison(3, "between", 10, 20)

# Blanks and nonblanks
data.autofilter.apply_comparison(2, "equal_to", None)
data.autofilter.apply_comparison(2, "not_equal_to", None)

# Clear one field or all fields
data.autofilter.clear(2)
data.autofilter.clear()
```

Value filters accept one or more strings, finite numbers, or booleans. Comparison filters accept `"between"`, `"not_between"`, `"equal_to"`, `"not_equal_to"`, `"greater_than"`, `"less_than"`, `"greater_than_or_equal"`, and `"less_than_or_equal"`. The between operators require two values; the other operators accept one.

Range AutoFilters are supported on Windows and with Office.js clients such as xlwings Server and xlwings Lite. Excel's macOS automation API doesn't support filtering an ordinary range, so use a table there. Table AutoFilters are supported on all three engines.

On Office.js clients, range AutoFilters require ExcelApi 1.14 and table AutoFilters require ExcelApi 1.2. Applying a range filter raises an error if the worksheet already has an AutoFilter on a different range. Clearing criteria leaves filter controls and sort state intact.

```{autoclass} xlwings.main.AutoFilter
:members:
```
