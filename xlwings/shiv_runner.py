"""This file is used by shiv in two contexts:
1. It provides the preamble command that is run when the file is started as a script (in __main__).
  shiv will run this script as a preamble each time a shived environment is started.
2. It provides the entry point main() that is run when the shived environment is started (after the preamble). This
  main function will:
- return the path to the shived environment (usually in USERPROFILE\.shiv\name_of_shived_pyz...) if run without arguments.
  This is used in the xlwings VBA code to build the path to the xlwings....dlls.
- execute the last argument given if at least one argument is given.
  This is used by xlwings to start the server.
"""
import sys
from pathlib import Path

# These variables are injected by shiv.bootstrap
site_packages: Path


def preamble():
    """The preamble will ensure that previously decompressed environments
    with the same name as the current environment are deleted from the .shiv folder.

    It will also move any DLL found in the 'site_packages\pywin32_system32' folder to the
    'site_packages\pwin32\lib' for pywin32 to find the DLLs (e.g. pywintypes38.dll)
    """
    import shutil

    # Get a handle of the current PYZ's site_packages directory
    current = site_packages.parent

    # The parent directory of the site_packages directory is our shiv cache
    cache_path = current.parent

    name, build_id = current.name.rsplit("_", 1)

    # remove old version of env
    for path in cache_path.iterdir():
        try:
            if path.name.startswith(f".{name}_") and not path.name.endswith(
                f"{build_id}_lock"
            ):
                path.unlink()
            if path.name.startswith(f"{name}_") and not path.name.endswith(build_id):
                shutil.rmtree(path)
        except PermissionError:
            pass

    # copy pywin32 dll from ...\site-packages\pywin32_system32
    # to ...\site-packages\win32\lib\pywintypes38.dll
    for dll in (site_packages / "pywin32_system32").glob("*.dll"):
        dll.rename(site_packages / "win32" / "lib" / dll.name)


def get_shiv_path():
    """Detect the path to the shiv folder"""
    for pth in sys.path:
        if ".shiv" in pth:
            pth = Path(pth)
            while not pth.parent.name == ".shiv":
                pth = pth.parent
            return pth
    return "path to shiv could not be found"


def main():
    if len(sys.argv) == 1:
        # if called without arguments, print the path to the .shiv subfolder of the environment
        print(get_shiv_path())
    else:
        # execute the last argument of the command with the python code from xlwings
        exec(sys.argv[-1])


if __name__ == "__main__":
    preamble()
