"""
Required Notice: Copyright (C) Zoomer Analytics GmbH.

xlwings PRO is dual-licensed under one of the following licenses:

* PolyForm Noncommercial License 1.0.0 (for noncommercial use):
  https://polyformproject.org/licenses/noncommercial/1.0.0
* xlwings PRO License (for commercial use):
  https://github.com/xlwings/xlwings/blob/main/LICENSE_PRO.txt

Commercial licenses can be purchased at https://www.xlwings.org
"""

from ...utils import fspath


class Image:
    """
    Use this class to provide images to either ``render_template()``.

    Arguments
    ---------

    filename : str or pathlib.Path object
        The file name or path
    """

    def __init__(self, filename):
        self.filename = fspath(filename)
