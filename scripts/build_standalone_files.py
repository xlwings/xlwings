import os
import re
import pywintypes
import xlwings as xw

this_dir = os.path.dirname(os.path.abspath(__file__))
par_dir = os.path.join(this_dir, os.path.pardir)

with open(os.path.join(par_dir, 'xlwings', '__init__.py')) as f:
    version = re.compile(r".*__version__ = '(.*?)'", re.S).match(f.read()).group(1)

addin_files = ['License.bas', 'Main.bas', 'Config.bas', 'Extensions.bas', 'Utils.bas']

with open("temp.bas", "w") as combined:
    for f in addin_files:
        with open(os.path.join(par_dir, "xlwings", "addin", f), "r") as component:
            combined.write(component.read())

with open("temp.bas", "r") as temp, open("xlwings.bas", "w") as xw_module:
    content = temp.read()
    content = content.replace("ActiveWorkbook", "ThisWorkbook")
    content = content.replace('Attribute VB_Name = "License"', "")
    content = content.replace("Attribute VB_Name", "\n'Attribute VB_Name")
    content = content.replace("Option Explicit", "")
    xw_module.seek(0, 0)
    xw_module.write('Attribute VB_Name = "xlwings"\n')
    xw_module.write("'Version: {}\n".format(version))

    xw_module.write(content)

os.remove("temp.bas")

# update standalone files
standalone_files = ['../xlwings/quickstart_standalone_mac.xlsm', '../xlwings/quickstart_standalone_win.xlsm']

for f in standalone_files:
    wb = xw.Book(os.path.abspath(f))
    try:
        wb.api.VBProject.VBComponents.Remove(wb.api.VBProject.VBComponents("xlwings"))
    except pywintypes.com_error:
        pass
    wb.api.VBProject.VBComponents.Import(os.path.abspath(os.path.join(this_dir, 'xlwings.bas')))
    wb.save()

for f in standalone_files:
    xw.Book(os.path.abspath(f)).close()