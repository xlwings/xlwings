import os
import re

# pythonnet
import clr
dll = os.path.abspath(os.path.join(os.getenv('APPVEYOR_BUILD_FOLDER', '..'), "aspose.cells", "lib", "net40", "Aspose.Cells.dll"))
clr.AddReference(dll)
from Aspose.Cells import Workbook, License


this_dir = os.path.dirname(os.path.abspath(__file__))
par_dir = os.path.join(this_dir, os.path.pardir)
addin_path = os.path.join(par_dir, 'xlwings', 'addin', 'xlwings.xlam')
standalone_win_path = os.path.join(par_dir, 'xlwings', 'quickstart_standalone_win.xlsm')
standalone_mac_path = os.path.join(par_dir, 'xlwings', 'quickstart_standalone_mac.xlsm')
xlwings_bas_path = os.path.join(par_dir, 'xlwings', 'xlwings.bas')
version = os.getenv('APPVEYOR_BUILD_VERSION', 'dev')

if os.getenv('ASPOSE_LICENSE'):
    license = License()
    license.SetLicense(os.path.abspath(os.path.join(this_dir, 'Aspose.Cells.lic')))


def set_version_strings(code):
    code = re.sub(r'XLWINGS_VERSION As String = ".*"',
                  'XLWINGS_VERSION As String = "{}"'.format(version),
                  code)
    code = code.replace("xlwings32.dll", "xlwings32-{}.dll".format(version))
    code = code.replace("xlwings64.dll", "xlwings64-{}.dll".format(version))
    return code


# Get vba modules from addin
addin_wb = Workbook(addin_path)
addin_modules = addin_wb.VbaProject.get_Modules()

# Update Main module in addin
main_code = addin_modules['Main'].get_Codes()
main_code = set_version_strings(main_code)
addin_modules['Main'].set_Codes(main_code)
addin_wb.Save(addin_path)

# Update standalone files with a single vba module containing the concatenated addin modules
standalone_code = ''
for m in ['License', 'Main', 'Config', 'Extensions', 'Utils']:
    standalone_code += addin_modules[m].get_Codes()

standalone_code = set_version_strings(standalone_code)
standalone_code = "'Version: {}\n".format(version) + standalone_code
standalone_code = standalone_code.replace("ActiveWorkbook", "ThisWorkbook")
standalone_code = standalone_code.replace('Attribute VB_Name = "License"', "")
standalone_code = standalone_code.replace("Attribute VB_Name", "\n'Attribute VB_Name")
standalone_code = standalone_code.replace("Option Explicit", "")

for path in [standalone_mac_path, standalone_win_path]:
    wb = Workbook(path)
    wb.VbaProject.get_Modules()['xlwings'].set_Codes(standalone_code)
    wb.Save(path)

# Save standalone as xlwings.bas to be included in python package
with open(xlwings_bas_path, 'w') as f:
    f.write('Attribute VB_Name = "xlwings"\n' + '\n'.join(standalone_code.splitlines()))

