# Updates all templates/examples with the latest VBA xlwings module

# Only runs on Windows and Excel must be set to "Trust access to the VBA project object model"
# under Options > Trust Center > Trust Center Settings > Macro Settings (in the case of Excel 2010)

import os
import xlwings as xw
from xlwings.constants import FileFormat

this_dir = os.path.dirname(os.path.abspath(__file__))
exclude_dirs = ['build', 'dist']

# Template
template_path = os.path.abspath(os.path.join(this_dir, os.pardir, 'xlwings', 'xlwings_template.xltm'))
workbook_paths = [template_path]

# Examples
root = os.path.abspath(os.path.join(this_dir, os.pardir))
for root, dirs, files in os.walk(root, topdown=True):
    for f in files:
        dirs[:] = [d for d in dirs if d not in exclude_dirs]
        if f.endswith(".xlsm") and not f == 'macro book.xlsm':
            workbook_paths.append((os.path.join(root, f)))

for path in workbook_paths:
    wb = xw.Book(path)
    wb.api.VBProject.VBComponents.Remove(wb.api.VBProject.VBComponents("xlwings"))
    wb.api.VBProject.VBComponents.Import(os.path.abspath(os.path.join(this_dir, os.pardir, 'xlwings', 'xlwings.bas')))
    if 'xlwings_template' in wb.fullname:
        # TODO: implement FileFormat in xlwings
        wb.api.Application.DisplayAlerts = False
        wb.api.SaveAs(template_path, FileFormat=FileFormat.xlOpenXMLTemplateMacroEnabled)
        wb.api.Application.DisplayAlerts = True
    else:
        wb.save()

for path in workbook_paths:
    xw.Book(path).close()



