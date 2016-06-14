# Updates all templates/examples with the latest VBA xlwings module

# Only runs on Windows and Excel must be set to "Trust access to the VBA project object model"
# under Options > Trust Center > Trust Center Settings > Macro Settings (in the case of Excel 2010)

import os
from xlwings import Book, FileFormat

this_dir = os.path.dirname(os.path.abspath(__file__))

# Template
template_path = os.path.abspath(os.path.join(this_dir, os.pardir, 'xlwings', 'xlwings_template.xltm'))
workbook_paths = [template_path]

# Examples
root = os.path.abspath(os.path.join(this_dir, os.pardir))
for root, dirs, files in os.walk(root):
    for f in files:
        if f.endswith(".xlsm") and not f == 'macro book.xlsm':
            workbook_paths.append((os.path.join(root, f)))

for path in workbook_paths:
    wb = Book(path)
    wb.xl_workbook.VBProject.VBComponents.Remove(wb.xl_workbook.VBProject.VBComponents("xlwings"))
    wb.xl_workbook.VBProject.VBComponents.Import(os.path.abspath(os.path.join(this_dir, os.pardir, 'xlwings', 'xlwings.bas')))
    if 'xlwings_template' in wb.fullname:
        # TODO: implement FileFormat in xlwings
        wb.xl_workbook.Application.DisplayAlerts = False
        wb.xl_workbook.SaveAs(template_path, FileFormat=FileFormat.xlOpenXMLTemplateMacroEnabled)
        wb.xl_workbook.Application.DisplayAlerts = True
    else:
        wb.save()
    wb.close()



