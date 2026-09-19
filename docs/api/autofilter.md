# AutoFilter

Use {attr}`Range.autofilter <xlwings.Range.autofilter>` or {attr}`Table.autofilter <xlwings.main.Table.autofilter>` to apply and clear filters without changing cell values, formulas, formats, or table identity. Fields are one-based positions relative to the range or table.

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

## `AutoFilter.apply_values(field, values)`

Filters a field to rows matching any supplied value.

- `field`: One-based column position relative to the range or table.
- `values`: A nonempty sequence of exact strings, finite numbers, or booleans to include.

Raises `TypeError` if `field` isn't an integer or `values` isn't a sequence of supported scalar values. Raises `ValueError` if the field is outside the target, the sequence is empty, or a number isn't finite.

## `AutoFilter.apply_comparison(field, operator, value1, value2=None)`

Filters a field using a comparison. `operator` accepts `"between"`, `"not_between"`, `"equal_to"`, `"not_equal_to"`, `"greater_than"`, `"less_than"`, `"greater_than_or_equal"`, or `"less_than_or_equal"`.

- `field`: One-based column position relative to the range or table.
- `operator`: The comparison to apply.
- `value1`: A string, finite number, boolean, or `None`. Use `None` with `"equal_to"` for blanks or with `"not_equal_to"` for nonblanks.
- `value2`: The required upper bound for `"between"` and `"not_between"`; omit it for every other operator.

Raises `TypeError` for unsupported value types. Raises `ValueError` for an invalid field, operator, operand combination, or nonfinite number.

## `AutoFilter.clear(field=None)`

Clears the criteria for one field. Omitting `field` clears criteria for every field while retaining the filter controls and sort state.

- `field`: Optional one-based column position relative to the range or table.
