import os
import clr

dll = os.path.abspath(os.path.join(os.environ['APPVEYOR_BUILD_FOLDER'], "aspose.cells", "lib", "net40", "Aspose.Cells.dll"))
clr.AddReference(dll)

from Aspose.Cells import Workbook

addin_path = os.path.join(os.path.join(os.environ['APPVEYOR_BUILD_FOLDER'], 'xlwings', 'addin', 'xlwings.xlam'))
workbook = Workbook(addin_path)

module = workbook.VbaProject.get_Modules()['Main']
code = module.get_Codes()
LINEBREAK = '\r\n'
code = code.split(LINEBREAK)
for i, line in enumerate(code):
    if line.startswith('Public Const XLWINGS_VERSION As String'):
        break

code[i] = 'Public Const XLWINGS_VERSION As String = "{}"'.format(os.environ['APPVEYOR_BUILD_VERSION'])
code = LINEBREAK.join(code)
module.set_Codes(code)
workbook.Save(addin_path)
