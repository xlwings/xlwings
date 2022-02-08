#!/usr/bin/env python3
import sys
from pathlib import Path

# These variables are injected by shiv.bootstrap
site_packages: Path


def preamble():
    import shutil

    # Get a handle of the current PYZ's site_packages directory
    current = site_packages.parent

    # The parent directory of the site_packages directory is our shiv cache
    cache_path = current.parent

    name, build_id = current.name.rsplit("_", 1)

    # remove old version of env
    for path in cache_path.iterdir():
        try:
            if path.name.startswith(f".{name}_") and not path.name.endswith(f"{build_id}_lock"):
                path.unlink()
            if path.name.startswith(f"{name}_") and not path.name.endswith(build_id):
                shutil.rmtree(path)
        except PermissionError:
            pass

    # copy pywin32 dll from ...\site-packages\pywin32_system32 to ...\site-packages\win32\lib\pywintypes38.dll
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
    # handle the case of scripted called via shebang
    if len(sys.argv) > 1 and (sys.argv[0] in sys.argv[1]):
        # script path has been duplicated due to the shebang call
        del sys.argv[1]

    if len(sys.argv) == 1:
        # if called without arguments, print the path to the .shiv subfolder with the environment
        print(get_shiv_path())
    else:

        # execute the last argument of the command with the python code from xlwings
        exec(sys.argv[-1])


if __name__ == "__main__":
    preamble()
