"""
Required Notice: Copyright (C) Zoomer Analytics GmbH.

xlwings PRO is dual-licensed under one of the following licenses:

* PolyForm Noncommercial License 1.0.0 (for noncommercial use):
  https://polyformproject.org/licenses/noncommercial/1.0.0
* xlwings PRO License (for commercial use):
  https://github.com/xlwings/xlwings/blob/main/LICENSE_PRO.txt

Commercial licenses can be purchased at https://www.xlwings.org
"""

import logging

from ... import XlwingsError

try:
    import pdfrw
except ImportError:
    pdfrw = None


def print_on_layout(report_path, layout_path):
    if not pdfrw:
        raise XlwingsError(
            "Couldn't find the 'pdfrw' package, which is required when using 'layout'."
        )
    report = pdfrw.PdfReader(report_path)
    layout = pdfrw.PdfReader(layout_path)

    for ix, page in enumerate(report.pages):
        if len(layout.pages) == 1:
            # Same layout for whole report
            layout_page_ix = 0
        elif len(report.pages) == len(layout.pages):
            # Every report page has a corresponding page in the layout
            layout_page_ix = ix
        else:
            raise XlwingsError(
                "The layout PDF must either be a single page or have the "
                f"same number of pages as the report (report: {len(report.pages)}, "
                f"layout: {len(layout.pages)})"
            )
        merge = pdfrw.PageMerge().add(layout.pages[layout_page_ix])[0]
        pdfrw.PageMerge(page).add(merge, prepend=True).render()
    # Changing log level as the exported PDFs from Excel aren't fully compliant and
    # would log the following:
    # [WARNING] tokens.py:221 Did not find PDF object (12, 0) (...)
    logging.getLogger("pdfrw").setLevel(logging.CRITICAL)
    pdfrw.PdfWriter(report_path, trailer=report).write()
