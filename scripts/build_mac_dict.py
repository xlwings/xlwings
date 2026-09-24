from pathlib import Path

from appscript.terminology import buildtablesforsdef, dumptables

excel = Path("/Applications/Microsoft Excel.app")
sdef = excel / "Contents/Resources/Excel.sdef"

tables = buildtablesforsdef(sdef.read_bytes())
dumptables(tables, str(excel), "xlwings/mac_dict_new.py")
