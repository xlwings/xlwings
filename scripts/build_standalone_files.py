import os
import re
from oletools.olevba3 import VBA_Parser
import pywintypes
import xlwings as xw

this_dir = os.path.dirname(os.path.abspath(__file__))
par_dir = os.path.join(this_dir, os.path.pardir)

version = 'test-version-1'


def parse(workbook_path):
    vba_path = workbook_path + '.vba'
    vba_parser = VBA_Parser(workbook_path)
    vba_modules = vba_parser.extract_all_macros() if vba_parser.detect_vba_macros() else []

    for _, _, _, content in vba_modules:
        decoded_content = content.decode('latin-1')
        lines = []
        if '\r\n' in decoded_content:
            lines = decoded_content.split('\r\n')
        else:
            lines = decoded_content.split('\n')
        if lines:
            name = lines[0].replace('Attribute VB_Name = ', '').strip('"')
            content = [lines[0]] + [line for line in lines[1:] if not (
                line.startswith('Attribute') and 'VB_' in line)]
            if content and content[-1] == '':
                content.pop(len(content)-1)
                non_empty_lines_of_code = len([c for c in content if c])
                if non_empty_lines_of_code > 1:
                    with open(os.path.join(name + '.bas'), 'w') as f:
                        f.write('\n'.join(content))


parse('../xlwings/addin/xlwings.xlam')


addin_files = ['License.bas', 'Main.bas', 'Config.bas', 'Extensions.bas', 'Utils.bas']

# standalone module
with open("temp.bas", "w") as combined:
    for f in addin_files:
        with open(f, "r") as component:
            combined.write(component.read())

with open("temp.bas", "r") as temp, open("xlwings.bas", "w") as xw_module:
    content = temp.read()
    content = content.replace("ActiveWorkbook", "ThisWorkbook")
    content = content.replace('Attribute VB_Name = "License"', "")
    content = content.replace("Attribute VB_Name", "\n'Attribute VB_Name")
    content = content.replace("Option Explicit", "")
    content = content.replace("xlwings32.dll", "xlwings32-{}.dll".format(version))
    content = content.replace("xlwings64.dll", "xlwings64-{}.dll".format(version))
    xw_module.seek(0, 0)
    xw_module.write('Attribute VB_Name = "xlwings"\n')
    xw_module.write("'Version: {}\n".format(version))

    xw_module.write(content)

os.remove("temp.bas")

# main.bas for add-in
os.replace('Main.bas', 'MainTemp.bas')
with open("MainTemp.bas", "r") as temp, open("Main.bas", "w") as xw_module:
    content = temp.read()
    content = content.replace("xlwings32.dll", "xlwings32-{}.dll".format(version))
    content = content.replace("xlwings64.dll", "xlwings64-{}.dll".format(version))
    content = re.sub(r'XLWINGS_VERSION As String = ".*"','XLWINGS_VERSION As String = "{}"'.format(version), content)
    xw_module.write(content)

os.remove("MainTemp.bas")

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

# addin

for f in ['../xlwings/addin/xlwings.xlam']:
    wb = xw.Book(os.path.abspath(f))
    try:
        wb.api.VBProject.VBComponents.Remove(wb.api.VBProject.VBComponents("Main"))
    except pywintypes.com_error:
        pass
    wb.api.VBProject.VBComponents.Import(os.path.abspath(os.path.join(this_dir, 'Main.bas')))
    wb.save()

for f in standalone_files:
    xw.Book(os.path.abspath(f)).close()