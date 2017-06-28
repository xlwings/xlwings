import os

files = ['Main.bas', 'Config.bas', 'Extensions.bas', 'Utils.bas']  # Main needs to come first

par_dir = os.path.dirname(os.path.abspath(os.path.join(__file__, os.path.pardir)))

with open("temp.bas", "w") as combined:
    for f in files:
        with open(os.path.join(par_dir, "addin", f), "r") as component:
            combined.write(component.read())

with open("temp.bas", "r") as temp, open("xlwings.bas", "w") as xw_module:
    content = temp.read()
    content = content.replace("ActiveWorkbook", "ThisWorkbook")
    content = content.replace("Attribute VB_Name", "'Attribute VB_Name")
    xw_module.seek(0, 0)
    xw_module.write('Attribute VB_Name = "xlwings"\n')

    xw_module.write(content)

os.remove("temp.bas")
