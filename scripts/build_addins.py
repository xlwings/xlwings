import os
from xlwings import Book
from xlwings.constants import FileFormat

"""
TODO: might make sense to refactor VBA module like this instead of replacing ThisWorkbook with ActiveWorkbook:
If ThisWorkbook.IsAddin Then
 ' this code runs when the workbook is an add-in
End If
"""

this_dir = os.path.dirname(os.path.abspath(__file__))
par_dir = os.path.dirname(os.path.abspath(os.path.join(__file__, os.path.pardir)))


def build_addins():
    # transform code for addin use
    with open(os.path.join(par_dir, "xlwings", "xlwings.bas"), "r") as vba_module, \
         open(os.path.join(this_dir, "xlwings_addin.bas"), "w") as vba_addin:
        content = vba_module.read().replace("ThisWorkbook", "ActiveWorkbook")
        content = content.replace('Attribute VB_Name = "xlwings"', 'Attribute VB_Name = "xlwings_addin"')
        vba_addin.write(content)

    # create addin workbook
    wb = Book()

    # remove unneeded sheets
    for sh in list(wb.sheets.api)[1:]:
        sh.Delete()

    # rename vbproject
    wb.api.VBProject.Name = "xlwings"
    
    # import modules
    wb.api.VBProject.VBComponents.Import(os.path.join(this_dir, "xlwings_addin.bas"))
    
    # save to xla and xlam
    wb.api.IsAddin = True
    wb.app.display_alerts = False
    # wb.api.SaveAs(os.path.join(this_dir, "xlwings.xla"), FileFormat.xlAddIn)
    wb.api.SaveAs(os.path.join(this_dir, "xlwings.xlam"), FileFormat.xlOpenXMLAddIn)
    wb.app.display_alerts = True

    # clean up
    wb.close()
    os.remove(os.path.join(this_dir, 'xlwings_addin.bas'))

if __name__ == '__main__':
    build_addins()
