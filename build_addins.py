import os
from xlwings import Workbook


def build_addins():
    # transform code for addin use
    with open(r"xlwings\xlwings.bas", "r") as original, open(r".\xlwings.bas", "w") as for_addin:
        for_addin.write(original.read().replace("ThisWorkbook", "ActiveWorkbook"))

    # create addin workbook
    wb = Workbook().xl_workbook
    wb.Application.Visible = True

    # remove unneeded sheets
    for sh in list(wb.Sheets)[1:]:
        sh.Delete()

    # rename vbproject
    wb.VBProject.Name = "xlwings_addin"
    
    # import modules
    fld = os.path.dirname(os.path.abspath(__file__))
    wb.VBProject.VBComponents.Import(os.path.join(fld, r"xlwings.bas"))
    
    # save to xla and xlam
    wb.IsAddin = True
    wb.SaveAs(os.path.join(fld, "xlwings_addin.xla"), 18)
    wb.SaveAs(os.path.join(fld, "xlwings_addin.xlam"), 55)


if __name__ == '__main__':
    build_addins()
