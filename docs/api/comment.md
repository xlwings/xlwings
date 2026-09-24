# Comment

A modern threaded comment is separate from a cell [](note.md). Use {meth}`Range.add_comment <xlwings.Range.add_comment>` to create one. On xlwings Lite, read an existing comment with `await cell.get_comment()`; fetch its text and metadata with the asynchronous methods below. Comment and reply mutations are queued, so call `await book.flush()` before reading a mutation back in the same script.

Comment creation, text changes, deletion, and replies are supported on Windows and Office.js. Resolving and reopening require xlwings Lite with ExcelApi 1.11. Excel for Mac's AppleScript interface does not expose threaded comments. Comment and reply reads in xlwings Lite require ExcelApi 1.10. A comment converted from a note may have no creation date.

```{eval-rst}
.. autoclass:: xlwings.main.Comment
    :members:
```
