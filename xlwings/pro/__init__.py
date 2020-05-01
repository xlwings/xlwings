"""(c) 2020-present Zoomer Analytics GmbH

Note that the code in this directory is licensed under a commercial license and must be used with a valid license key.

You will find the license under LICENSE_PRO.txt in the root directory or under
https://github.com/xlwings/xlwings/blob/master/LICENSE_PRO.txt
"""

from .utils import LicenseHandler
from .embedded_code import runpython_embedded_code

LicenseHandler.validate_license('pro')
