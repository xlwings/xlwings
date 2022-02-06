import os
import re
import zipfile
import tempfile
import shutil

# pythonnet
import clr

dll = os.path.abspath(
    os.path.join(
        os.environ["GITHUB_WORKSPACE"], "aspose", "lib", "net40", "Aspose.Cells.dll"
    )
)
clr.AddReference(dll)
from Aspose.Cells import Workbook, License


this_dir = os.path.dirname(os.path.abspath(__file__))
par_dir = os.path.join(this_dir, os.path.pardir)
addin_path = os.path.join(par_dir, "xlwings", "addin", "xlwings.xlam")
addin_unprotected_path = os.path.join(
    par_dir, "xlwings", "addin", "xlwings_unprotected.xlam"
)
standalone_path = os.path.join(par_dir, "xlwings", "quickstart_standalone.xlsm")
myaddin_ribbon_path = os.path.join(par_dir, "xlwings", "quickstart_addin_ribbon.xlam")
myaddin_path = os.path.join(par_dir, "xlwings", "quickstart_addin.xlam")
xlwings_bas_path = os.path.join(par_dir, "xlwings", "xlwings.bas")

# Version string
if os.environ["GITHUB_REF"].startswith("refs/tags"):
    version_string = os.environ["GITHUB_REF"][10:]
else:
    version_string = os.environ["GITHUB_SHA"][:7]

# Rename dlls and applescript file
for i in ["32", "64"]:
    os.rename(
        os.path.join(os.environ["GITHUB_WORKSPACE"], "xlwings{0}.dll".format(i)),
        os.path.join(
            os.environ["GITHUB_WORKSPACE"],
            "xlwings{0}-{1}.dll".format(i, version_string),
        ),
    )

os.rename(
    os.path.join(os.environ["GITHUB_WORKSPACE"], "xlwings", "xlwings-dev.applescript"),
    os.path.join(
        os.environ["GITHUB_WORKSPACE"],
        "xlwings",
        f"xlwings-{version_string}.applescript",
    ),
)
# Stamp version
version_file = os.path.join(os.environ["GITHUB_WORKSPACE"], "xlwings", "__init__.py")
cli_file = os.path.join(os.environ["GITHUB_WORKSPACE"], "xlwings", "cli.py")
for source_file in [version_file, cli_file]:
    with open(source_file, "r") as f:
        content = f.read()
    content = content.replace("dev", version_string)
    with open(source_file, "w") as f:
        f.write(content)

# License handler
lh = os.path.join(os.environ["GITHUB_WORKSPACE"], "xlwings", "pro", "utils.py")
with open(lh, "r") as f:
    content = f.read()
content = content.replace(
    "os.getenv('XLWINGS_LICENSE_KEY_SECRET')",
    "'" + os.environ["XLWINGS_LICENSE_KEY_SECRET"] + "'",
)
with open(lh, "w") as f:
    f.write(content)

# Aspose license
if os.getenv("ASPOSE_LICENSE"):
    lic_file = os.path.abspath(
        os.path.join(os.environ["GITHUB_WORKSPACE"], "aspose", "Aspose.Cells.lic")
    )
    with open(lic_file, "w") as f:
        f.write(os.environ["ASPOSE_LICENSE"])
    license = License()
    license.SetLicense(lic_file)


def set_version_strings(code):
    code = re.sub(
        r'XLWINGS_VERSION As String = ".*"',
        'XLWINGS_VERSION As String = "{}"'.format(version_string),
        code,
    )
    code = code.replace("xlwings32-dev.dll", "xlwings32-{}.dll".format(version_string))
    code = code.replace("xlwings64-dev.dll", "xlwings64-{}.dll".format(version_string))
    return code


def produce_single_module(addin_modules, custom_addin=False):
    # Read out modules
    vba_module_names = ["License", "Main", "Config", "Extensions", "Utils"]
    if custom_addin:
        vba_module_names.pop(vba_module_names.index("Extensions"))
    standalone_code = ""
    for name in vba_module_names:
        standalone_code += addin_modules[name].get_Codes()

    standalone_code = set_version_strings(standalone_code)
    standalone_code = "'Version: {}\n".format(version_string) + standalone_code
    if custom_addin:
        standalone_code = standalone_code.replace(
            'Public Const PROJECT_NAME As String = "xlwings"',
            'Public Const PROJECT_NAME As String = "myaddin"',
        )
    else:
        # TODO: handle this in the VBA code for standalone modules, too
        standalone_code = standalone_code.replace("ActiveWorkbook", "ThisWorkbook")
        standalone_code = standalone_code.replace("ActiveDocument", "ThisDocument")
    standalone_code = standalone_code.replace('Attribute VB_Name = "License"', "")
    standalone_code = standalone_code.replace(
        "Attribute VB_Name", "\n'Attribute VB_Name"
    )
    standalone_code = standalone_code.replace("Option Explicit", "")
    standalone_code = standalone_code.replace(
        """#Const App = "Microsoft Excel" 'Adjust when using outside of Excel""", ""
    )
    # Re-add the Compiler Constant
    standalone_code = (
        'Attribute VB_Name = "xlwings"\n'
        + """#Const App = "Microsoft Excel" 'Adjust when using outside of Excel\n"""
        + "\n".join(standalone_code.splitlines())
    )
    return standalone_code


# Get vba modules from addin
addin_wb = Workbook(addin_path)
addin_modules = addin_wb.VbaProject.get_Modules()

# Update Main module in xlwings add-in
main_code = addin_modules["Main"].get_Codes()
main_code = set_version_strings(main_code)
addin_modules["Main"].set_Codes(main_code)
addin_wb.Save(addin_path)

# Create an unprotected copy of the xlwings add-in (without password)
addin_unprotected_wb = Workbook(addin_path)
addin_unprotected_wb.VbaProject.Protect(False, None)
addin_unprotected_wb.Save(addin_unprotected_path)

# Save standalone module
standalone_code = produce_single_module(addin_modules, custom_addin=False)
wb = Workbook(standalone_path)
wb.VbaProject.get_Modules()["xlwings"].set_Codes(standalone_code)
wb.Save(standalone_path)

# Custom add-in
standalone_code_addin = produce_single_module(addin_modules, custom_addin=True)

for path in [myaddin_path, myaddin_ribbon_path]:
    wb = Workbook(path)
    wb.VbaProject.get_Modules()["xlwings"].set_Codes(standalone_code_addin)
    wb.Save(path)

# Save standalone as xlwings.bas to be included in python package
with open(xlwings_bas_path, "w") as f:
    f.write(standalone_code)


# Hack the _rels/.rels file in the add-in so the ribbon also shows up in Excel 2007
def update_zip(zipname, filename, data):
    # generate a temp file
    tmpfd, tmpname = tempfile.mkstemp(dir=os.path.dirname(zipname))
    os.close(tmpfd)

    # create a temp copy of the archive without filename
    with zipfile.ZipFile(zipname, "r") as zin:
        with zipfile.ZipFile(tmpname, "w") as zout:
            for item in zin.infolist():
                if item.filename != filename:
                    zout.writestr(item, zin.read(item.filename))

    # replace with the temp archive
    os.remove(zipname)
    os.rename(tmpname, zipname)

    # now add filename with its new data
    with zipfile.ZipFile(zipname, mode="a", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr(filename, data)


content = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/><Relationship Id="R09696ac1de4341b9" Type="http://schemas.microsoft.com/office/2006/relationships/ui/extensibility" Target="customUI/customUI.xml"/></Relationships>'
update_zip(addin_path, "_rels/.rels", content)

# Copy add-in to dist folder so it gets uploaded to artifacts
os.makedirs(os.path.join(os.environ["GITHUB_WORKSPACE"], "dist"), exist_ok=True)
shutil.copyfile(
    addin_path, os.path.join(os.environ["GITHUB_WORKSPACE"], "dist", "xlwings.xlam")
)

# Handle version stamp in JavaScript modules
for js in [
    os.path.join(par_dir, "xlwings", "js", "xlwings.ts"),
    os.path.join(par_dir, "xlwings", "js", "xlwings.js"),
]:
    with open(js, "r") as f:
        content = (
            f.read()
            .replace("* xlwings dev", f"* xlwings {version_string}")
            .replace('] = "dev";', f'] = "{version_string}";')
        )
    with open(js, "w") as f:
        f.write(content)
