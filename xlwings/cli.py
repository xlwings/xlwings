import os
import sys
import shutil
import argparse
import hashlib
import socket
import json
import tempfile
import subprocess
from pathlib import Path
from keyword import iskeyword

import xlwings as xw


# Directories/paths
this_dir = os.path.dirname(os.path.realpath(__file__))


def exit_on_mac():
    if sys.platform.startswith("darwin"):
        sys.exit("This command is currently only supported on Windows.")


def get_addin_dir():
    # The call to startup_path creates the XLSTART folder if it doesn't exist yet
    if xw.apps:
        return xw.apps.active.startup_path
    else:
        with xw.App(visible=False) as app:
            startup_path = app.startup_path
        return startup_path


def addin_install(args):
    xlwings_addin_target_path = os.path.join(get_addin_dir(), "xlwings.xlam")
    addin_name = "xlwings_unprotected.xlam" if args.unprotected else "xlwings.xlam"
    try:
        if args.file:
            custom_addin_source_path = os.path.abspath(args.file)
            shutil.copyfile(
                custom_addin_source_path,
                os.path.join(
                    get_addin_dir(), os.path.basename(custom_addin_source_path)
                ),
            )
            print("Successfully installed the add-in! Please restart Excel.")
        elif args.dir:
            for f in Path(args.dir).resolve().glob("[!~$]*.xl*"):
                shutil.copyfile(f, os.path.join(get_addin_dir(), f.name))
        else:
            shutil.copyfile(
                os.path.join(this_dir, "addin", addin_name), xlwings_addin_target_path
            )
            print("Successfully installed the xlwings add-in! Please restart Excel.")
        if sys.platform.startswith("darwin"):
            runpython_install(None)
        if not args.file:
            config_create(None)
    except IOError as e:
        if e.args[0] == 13:
            print(
                "Error: Failed to install the add-in: If Excel is running, "
                "quit Excel and try again."
            )
        else:
            print(repr(e))
    except Exception as e:
        print(repr(e))


def addin_remove(args):
    if args.file:
        addin_name = os.path.basename(args.file)
    else:
        addin_name = "xlwings.xlam"
    addin_path = os.path.join(get_addin_dir(), addin_name)
    try:
        os.remove(addin_path)
        print("Successfully removed the add-in!")
    except (WindowsError, PermissionError) as e:
        if e.args[0] in (13, 32):
            print(
                "Error: Failed to remove the add-in: If Excel is running, "
                "quit Excel and try again. "
                "You can also delete it manually from {0}".format(addin_path)
            )
        elif e.args[0] == 2:
            print(
                "Error: Could not remove the add-in. "
                "The add-in doesn't seem to be installed."
            )
        else:
            print(repr(e))
    except Exception as e:
        print(repr(e))


def addin_status(args):
    if args.file:
        addin_name = os.path.basename(args.file)
    else:
        addin_name = "xlwings.xlam"
    addin_path = os.path.join(get_addin_dir(), addin_name)
    if os.path.isfile(addin_path):
        print("The add-in is installed at {}".format(addin_path))
        print('Use "xlwings addin remove" to uninstall it.')
    else:
        print("The add-in is not installed.")
        print('"xlwings addin install" will install it at: {}'.format(addin_path))


def quickstart(args):
    project_name = args.project_name
    if not project_name.isidentifier() or iskeyword(project_name):
        sys.exit(
            "Error: You must choose a project name that works as Python module, "
            "i.e., it must only use letters, underscores and numbers and must not "
            "start with a number. Note that you *can* rename your Excel file "
            "manually, if you also adjust your RunPython VBA function "
            "accordingly."
        )
    cwd = os.getcwd()

    if args.fastapi:
        # Raises an error on its own if the dir already exists
        shutil.copytree(
            Path(this_dir) / "quickstart_fastapi",
            Path(cwd) / project_name,
            ignore=shutil.ignore_patterns("__pycache__"),
        )
        sys.exit(0)

    # Project dir
    project_path = os.path.join(cwd, project_name)
    if not os.path.exists(project_path):
        os.makedirs(project_path)
    else:
        sys.exit("Error: Directory already exists.")

    # Python file
    with open(os.path.join(project_path, project_name + ".py"), "w") as python_module:
        python_module.write("import xlwings as xw\n\n\n")
        python_module.write("def main():\n")
        python_module.write("    wb = xw.Book.caller()\n")
        python_module.write("    sheet = wb.sheets[0]\n")
        python_module.write('    if sheet["A1"].value == "Hello xlwings!":\n')
        python_module.write('        sheet["A1"].value = "Bye xlwings!"\n')
        python_module.write("    else:\n")
        python_module.write('        sheet["A1"].value = "Hello xlwings!"\n\n\n')
        if sys.platform.startswith("win"):
            python_module.write("@xw.func\n")
            python_module.write("def hello(name):\n")
            python_module.write('    return f"Hello {name}!"\n\n\n')
        python_module.write('if __name__ == "__main__":\n')
        python_module.write(
            '    xw.Book("{0}.xlsm").set_mock_caller()\n'.format(project_name)
        )
        python_module.write("    main()\n")

    # Excel file
    if args.standalone:
        source_file = os.path.join(this_dir, "quickstart_standalone.xlsm")
    elif args.addin and args.ribbon:
        source_file = os.path.join(this_dir, "quickstart_addin_ribbon.xlam")
    elif args.addin:
        source_file = os.path.join(this_dir, "quickstart_addin.xlam")
    else:
        source_file = os.path.join(this_dir, "quickstart.xlsm")

    shutil.copyfile(
        source_file,
        os.path.join(project_path, project_name + os.path.splitext(source_file)[1]),
    )


def runpython_install(args):
    destination_dir = (
        os.path.expanduser("~") + "/Library/Application Scripts/com.microsoft.Excel"
    )
    if not os.path.exists(destination_dir):
        os.makedirs(destination_dir)
    shutil.copy(
        os.path.join(this_dir, f"xlwings-{xw.__version__}.applescript"), destination_dir
    )
    print("Successfully enabled RunPython!")


def restapi_run(args):
    import subprocess

    try:
        import flask
    except ImportError:
        sys.exit("To use the xlwings REST API server, you need Flask>=1.0.0 installed.")
    host = args.host
    port = args.port

    os.environ["FLASK_APP"] = "xlwings.rest.api"
    subprocess.check_call(["flask", "run", "--host", host, "--port", port])


def license_update(args):
    """license handler for xlwings PRO"""
    key = args.key
    if not key:
        sys.exit(
            "Please provide a license key via the -k/--key option. "
            "For example: xlwings license update -k MY_KEY"
        )
    license_kv = '"LICENSE_KEY","{0}"\n'.format(key)
    # Update xlwings.conf
    new_config = []
    if os.path.exists(xw.USER_CONFIG_FILE):
        with open(xw.USER_CONFIG_FILE, "r") as f:
            config = f.readlines()
        for line in config:
            # Remove existing license key and empty lines
            if line.split(",")[0] == '"LICENSE_KEY"' or line in ("\r\n", "\n"):
                pass
            else:
                new_config.append(line)
        new_config.append(license_kv)
    else:
        new_config = [license_kv]
    if not os.path.exists(os.path.dirname(xw.USER_CONFIG_FILE)):
        os.makedirs(os.path.dirname(xw.USER_CONFIG_FILE))
    with open(xw.USER_CONFIG_FILE, "w") as f:
        f.writelines(new_config)
    print("Successfully updated license key.")


def license_deploy(args):
    from .pro import LicenseHandler

    print(LicenseHandler.create_deploy_key())


def get_conda_settings():
    conda_env = os.getenv("CONDA_DEFAULT_ENV")
    conda_exe = os.getenv("CONDA_EXE")

    if conda_env and conda_exe:
        # xlwings currently expects the path
        # without the trailing /bin/conda or \Scripts\conda.exe
        conda_path = os.path.sep.join(conda_exe.split(os.path.sep)[:-2])
        return conda_path, conda_env
    else:
        return None, None


def config_create(args):
    if args is None:
        force = False
    else:
        force = args.force
    os.makedirs(os.path.dirname(xw.USER_CONFIG_FILE), exist_ok=True)
    settings = []
    conda_path, conda_env = get_conda_settings()
    if conda_path and sys.platform.startswith("win"):
        settings.append('"CONDA PATH","{}"\n'.format(conda_path))
        settings.append('"CONDA ENV","{}"\n'.format(conda_env))
    else:
        extension = "MAC" if sys.platform.startswith("darwin") else "WIN"
        settings.append('"INTERPRETER_{}","{}"\n'.format(extension, sys.executable))
    if os.path.exists(xw.USER_CONFIG_FILE) and not force:
        print(
            "There is already an existing ~/.xlwings/xlwings.conf file. Run "
            "'xlwings config create --force' if you want to reset your configuration."
        )
    else:
        with open(xw.USER_CONFIG_FILE, "w") as f:
            f.writelines(settings)


def code_embed(args):
    """Import a specific file or all Python files of the Excel books' directory
    into the active Excel Book
    """
    wb = xw.books.active
    screen_updating = wb.app.screen_updating
    wb.app.screen_updating = False

    if args and args.file:
        import_dir = False
        source_files = [Path(args.file)]
    else:
        import_dir = True
        source_files = list(Path(wb.fullname).resolve().parent.glob("*.py"))

    if not source_files:
        print("WARNING: Couldn't find any Python files in the workbook's directory!")
    for source_file in source_files:
        with open(source_file, "r", encoding="utf-8") as f:
            content = []
            for line in f.read().splitlines():
                # Handle single-quote docstrings
                line = line.replace("'''", '"""')
                # Duplicate leading single quotes so Excel interprets them properly
                # This is required even if the cell is in Text format
                content.append(["'" + line if line.startswith("'") else line])

        if source_file.name not in [sheet.name for sheet in wb.sheets]:
            sheet = wb.sheets.add(source_file.name, after=wb.sheets[len(wb.sheets) - 1])
        else:
            sheet = wb.sheets[source_file.name]
        sheet.cells.clear_contents()
        sheet["A1"].resize(row_size=len(content)).number_format = "@"
        sheet["A1"].value = content
        sheet["A:A"].column_width = 65

    # Cleanup: remove sheets that don't exist anymore as source files
    if import_dir:
        source_file_names = set([path.name for path in source_files])
        source_sheet_names = set(
            [sheet.name for sheet in wb.sheets if sheet.name.endswith(".py")]
        )
        for sheet_name in source_sheet_names.difference(source_file_names):
            wb.sheets[sheet_name].delete()

    wb.app.screen_updating = screen_updating


def print_permission_json(scope):
    from .pro import dump_embedded_code

    assert scope in ["cwd", "book"]
    if scope == "cwd":
        source_files = Path(".").glob("*.py")
    else:
        tempdir = tempfile.TemporaryDirectory(prefix="xlwings-")
        source_files = Path(tempdir.name).glob("*.py")
        dump_embedded_code(xw.books.active, tempdir.name)

    payload = {"modules": []}
    for source_file in source_files:
        with open(source_file, "rb") as f:
            content = f.read()
        payload["modules"].append(
            {
                "file_name": source_file.name,
                "sha256": hashlib.sha256(content).hexdigest(),
                "machine_names": [socket.gethostname()],
            }
        )
    print(json.dumps(payload, indent=2))
    if scope == "book":
        tempdir.cleanup()


def permission_cwd(args):
    print_permission_json("cwd")


def permission_book(args):
    print_permission_json("book")


def copy_os(args):
    copy_js("ts")


def copy_gs(args):
    copy_js("js")


def copy_js(extension):
    try:
        from pandas.io import clipboard
    except ImportError:
        try:
            import pyperclip as clipboard
        except ImportError:
            sys.exit(
                'Please install either "pandas" or "pyperclip" to use the copy command.'
            )

    with open(Path(this_dir) / "js" / f"xlwings.{extension}", "r") as f:
        clipboard.copy(f.read())
        print("Successfully copied to clipboard.")


def release(args):
    from xlwings.utils import query_yes_no, read_user_config
    from xlwings.pro import LicenseHandler

    if sys.platform.startswith("darwin"):
        sys.exit(
            "This command is currently only supported on Windows. "
            "However, a released workbook will work on macOS, too."
        )

    installation_dir = Path(xw.__file__).resolve().parent

    if xw.apps:
        book = xw.apps.active.books.active
    else:
        sys.exit("Please open your Excel file first.")

    # Deploy Key
    try:
        deploy_key = LicenseHandler.create_deploy_key()
    except xw.LicenseError:
        print(
            "WARNING: Couldn't create a deploy key, "
            "using an expiring license key instead!"
        )
        deploy_key = read_user_config()["license_key"]

    # Sheet Config
    if "xlwings.conf" not in [i.name for i in book.sheets]:
        project_name = input("Name of your one-click installer? ")
        use_embedded_code = query_yes_no("Embed your Python code?")
        hide_config_sheet = query_yes_no("Hide the config sheet?")
        if use_embedded_code:
            hide_code_sheets = query_yes_no(
                "Hide the sheets with the embedded Python code?"
            )
        else:
            hide_code_sheets = False
        use_without_addin = query_yes_no(
            "Allow your tool to run without the xlwings add-in?"
        )
        print()
        if not query_yes_no(f'This will release "{book.name}", proceed?'):
            sys.exit()
        else:
            print()
            if "_xlwings.conf" in [sheet.name for sheet in book.sheets]:
                print("* Remove _xlwings.conf sheet")
                book.sheets["_xlwings.conf"].delete()
            active_sheet = book.sheets.active
            print("* Add xlwings.conf sheet")
            config_sheet = book.sheets.add(
                "xlwings.conf", after=book.sheets[len(book.sheets) - 1]
            )
            active_sheet.activate()  # preserve the currently active sheet
            config = {
                "Interpreter_Win": r"%LOCALAPPDATA%\{0}\python.exe".format(project_name)
                if project_name
                else None,
                "Interpreter_Mac": f"$HOME/{project_name}/bin/python"
                if project_name
                else None,
                "PYTHONPATH": None,
                "Conda Path": None,
                "Conda Env": None,
                "UDF Modules": None,
                "Debug UDFs": False,
                "Use UDF Server": False,
                "Show Console": False,
                "LICENSE_KEY": deploy_key,
                "RELEASE_EMBED_CODE": use_embedded_code,
                "RELEASE_HIDE_CONFIG_SHEET": hide_config_sheet,
                "RELEASE_HIDE_CODE_SHEETS": hide_code_sheets,
                "RELEASE_NO_ADDIN": use_without_addin,
            }
            config_sheet["A1"].value = config
            config_sheet["A:A"].autofit()
    else:
        print()
        if not query_yes_no(
            f'This will release "{book.name}" '
            f'according to the "xlwings.conf" sheet, proceed?'
        ):
            sys.exit()
        print()
        # Only update the deploy key
        config = xw.utils.read_config_sheet(book)
        print("* Update deploy key")
        config["LICENSE_KEY"] = deploy_key
        book.sheets["xlwings.conf"]["A1"].value = config

    # Remove Reference
    if config["RELEASE_NO_ADDIN"]:
        if "xlwings" in [i.Name for i in book.api.VBProject.References]:
            print("* Remove VBA Reference")
            ref = book.api.VBProject.References("xlwings")
            book.api.VBProject.References.Remove(ref)

        # Remove VBA modules/classes
        print("* Update VBA modules")
        if "xlwings" in [i.Name for i in book.api.VBProject.VBComponents]:
            book.api.VBProject.VBComponents.Remove(
                book.api.VBProject.VBComponents("xlwings")
            )

        if "Dictionary" in [i.Name for i in book.api.VBProject.VBComponents]:
            book.api.VBProject.VBComponents.Remove(
                book.api.VBProject.VBComponents("Dictionary")
            )

        # Import VBA modules/classes
        book.api.VBProject.VBComponents.Import(installation_dir / "xlwings.bas")
        book.api.VBProject.VBComponents.Import(installation_dir / "Dictionary.cls")

    # Embed code
    if config.get("RELEASE_EMBED_CODE"):
        print("* Embed Python code")
        code_embed(None)
    else:
        for sheet in book.sheets:
            if sheet.name.endswith(".py"):
                sheet.delete()

    # Hide sheets
    if config.get("RELEASE_HIDE_CONFIG_SHEET"):
        print("* Hide config sheet")
        book.sheets["xlwings.conf"].visible = False

    if config.get("RELEASE_HIDE_CODE_SHEETS"):
        print("* Hide Python sheets")
        for sheet in book.sheets:
            if sheet.name.endswith(".py"):
                sheet.visible = False
    print()
    print(
        "Checking for xlwings version compatibility "
        "between the one-click installer and the Excel file..."
    )
    if sys.platform.startswith("win") and config["Interpreter_Win"]:
        interpreter_path = os.path.expandvars(config["Interpreter_Win"])
    elif sys.platform.startswith("darwin") and config["Interpreter_Mac"]:
        interpreter_path = os.path.expandvars(config["Interpreter_Mac"])
    else:
        interpreter_path = None
    if interpreter_path and Path(interpreter_path).is_file():
        res = subprocess.run(
            [interpreter_path, "-c", "import xlwings;print(xlwings.__version__)"],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            encoding="utf-8",
        )
        xlwings_version_installer = res.stdout.strip()
        if xlwings_version_installer == xw.__version__:
            print(f'Successfully prepared "{book.name}" for release!')
        else:
            print(
                f"ERROR: You are running this command with xlwings {xw.__version__} "
                f"but your installer uses {xlwings_version_installer}!"
            )
    else:
        print(
            f"WARNING: Prepared '{book.name}' for release "
            "but couldn't verify the xlwings version!"
        )


def export_vba_modules(book, overwrite=False):
    # TODO: catch error when Trust Access to VBA Object model isn't enabled
    # TODO: raise error if editing while file hashes differ
    type_to_ext = {100: "cls", 1: "bas", 2: "cls", 3: "frm"}
    path_to_type = {}
    for vb_component in book.api.VBProject.VBComponents:
        file_path = (
            Path(book.fullname).resolve().parent
            / f"{vb_component.Name}.{type_to_ext[vb_component.Type]}"
        )
        path_to_type[str(file_path)] = vb_component.Type
        if vb_component.CodeModule.CountOfLines > 0:
            # Prevents cluttering everything with empty files if you have lots of sheets
            if overwrite:
                vb_component.Export(str(file_path))
            elif not file_path.exists():
                vb_component.Export(str(file_path))
            if vb_component.Type == 100:
                with open(file_path, "r") as f:
                    exported_code = f.readlines()
                with open(file_path, "w") as f:
                    f.writelines(exported_code[9:])
    return path_to_type


def vba_import(args):
    exit_on_mac()
    if args and args.file:
        book = xw.Book(args.file)
    else:
        if not xw.apps:
            sys.exit(
                "Your workbook must be open or you have to supply the --file argument."
            )
        else:
            book = xw.books.active
    for path in Path(book.fullname).resolve().parent.glob("*"):
        if path.suffix == ".bas":
            # This also imports frx, unlike in editing mode
            try:
                vb_component = book.api.VBProject.VBComponents(path.stem)
                book.api.VBProject.VBComponents.Remove(vb_component)
            except:
                pass
            book.api.VBProject.VBComponents.Import(path)
        elif path.suffix in (".cls", ".frm"):
            with open(path, "r") as f:
                vba_code = f.readlines()
            if vba_code:
                if vba_code[0].startswith("VERSION "):
                    try:
                        vb_component = book.api.VBProject.VBComponents(path.stem)
                        book.api.VBProject.VBComponents.Remove(vb_component)
                    except:
                        pass
                    book.api.VBProject.VBComponents.Import(path)
                else:
                    vb_component = book.api.VBProject.VBComponents(path.stem)
                    line_count = vb_component.CodeModule.CountOfLines
                    if line_count > 0:
                        vb_component.CodeModule.DeleteLines(1, line_count)
                    vb_component.CodeModule.AddFromString("".join(vba_code))


def vba_export(args):
    exit_on_mac()
    if args and args.file:
        book = xw.Book(args.file)
    else:
        if not xw.apps:
            sys.exit(
                "Your workbook must be open or you have to supply the --file argument."
            )
        else:
            book = xw.books.active
    export_vba_modules(book, overwrite=True)
    print(f"Successfully exported the VBA modules from {book.name}!")


def vba_edit(args):
    exit_on_mac()
    try:
        from watchgod import watch, RegExpWatcher, Change
    except ImportError:
        sys.exit(
            "Please install watchgod to use this functionality: pip install watchgod"
        )
    import pywintypes

    if args and args.file:
        book = xw.Book(args.file)
    else:
        if not xw.apps:
            sys.exit(
                "Your workbook must be open or you have to supply the --file argument."
            )
        else:
            book = xw.books.active

    path_to_type = export_vba_modules(book, overwrite=False)

    mode = "verbose" if args.verbose else "silent"

    print(f"NOTE: Deleting a VBA module here will also delete it in the VBA editor!")
    print(f"Watching for changes in {book.name} ({mode} mode)...(Hit Ctrl-C to stop)")

    for changes in watch(
        Path(book.fullname).resolve().parent,
        watcher_cls=RegExpWatcher,
        watcher_kwargs=dict(re_files=r"^.*(\.cls|\.frm|\.bas)$"),
        normal_sleep=400,
    ):
        for change_type, path in changes:
            module_name = os.path.splitext(os.path.basename(path))[0]
            vb_component = book.api.VBProject.VBComponents(module_name)
            if change_type == Change.modified:
                if path_to_type[path] in [100, 3]:
                    # ThisWorkbook/Sheet (100) don't accept VBComponents.Import
                    # and Forms (3) would overwrite the frx layout part
                    with open(path, "r") as f:
                        vba_code = f.readlines()
                    line_count = vb_component.CodeModule.CountOfLines
                    if line_count > 0:
                        vb_component.CodeModule.DeleteLines(1, line_count)
                    if path_to_type[path] == 100:
                        vb_component.CodeModule.AddFromString("".join(vba_code))
                    else:
                        # frm: ignore Attribute VB_ etc.
                        vb_component.CodeModule.AddFromString("".join(vba_code[15:]))
                    if args.verbose:
                        print(f"INFO: Updated module {module_name}.")
                else:
                    try:
                        book.api.VBProject.VBComponents.Remove(vb_component)
                        book.api.VBProject.VBComponents.Import(path)
                        if args.verbose:
                            print(f"INFO: Updated module {module_name}.")
                    except pywintypes.com_error:
                        print(
                            f"ERROR: Couldn't update module {module_name}. "
                            f"Please update changes manually."
                        )
            elif change_type == Change.deleted:
                try:
                    book.api.VBProject.VBComponents.Remove(vb_component)
                except pywintypes.com_error:
                    print(
                        f"ERROR: Couldn't delete module {module_name}. "
                        f"Please delete it manually."
                    )
            elif change_type == Change.added:
                print(
                    f"ERROR: Couldn't add {module_name} as this isn't supported. "
                    "Please add new files via the VBA Editor."
                )
            book.save()


def main():
    print("xlwings version: " + "dev")
    parser = argparse.ArgumentParser()
    subparsers = parser.add_subparsers(dest="command")
    subparsers.required = True

    # Add-in
    addin_parser = subparsers.add_parser(
        "addin",
        help='Run "xlwings addin install" to install the Excel add-in '
        '(will be copied to the XLSTART folder). Instead of "install" you can '
        'also use "update", "remove" or "status". Note that this command '
        'may take a while. Use the "--unprotected" flag to install the '
        "add-in without password protection. You can install your custom add-in "
        "by providing the name or path via the --file flag, "
        'e.g. "xlwings add-in install --file custom.xlam or copy all Excel '
        "files in a directory to the XLSTART folder by providing the path "
        'via the --dir flag."',
    )
    addin_subparsers = addin_parser.add_subparsers(dest="subcommand")
    addin_subparsers.required = True

    file_arg_help = "The name or path of a custom add-in."
    dir_arg_help = (
        "The path of a directory whose Excel files you want to copy to XLSTART."
    )

    addin_install_parser = addin_subparsers.add_parser("install")
    addin_install_parser.add_argument(
        "-u",
        "--unprotected",
        action="store_true",
        help="Install the add-in without the password protection.",
    )
    addin_install_parser.add_argument("-f", "--file", default=None, help=file_arg_help)
    addin_install_parser.add_argument("-d", "--dir", default=None, help=dir_arg_help)
    addin_install_parser.set_defaults(func=addin_install)

    addin_update_parser = addin_subparsers.add_parser("update")
    addin_update_parser.add_argument("-f", "--file", default=None, help=file_arg_help)
    addin_update_parser.add_argument("-d", "--dir", default=None, help=dir_arg_help)
    addin_update_parser.set_defaults(func=addin_install)

    addin_upgrade_parser = addin_subparsers.add_parser("upgrade")
    addin_upgrade_parser.add_argument("-f", "--file", default=None, help=file_arg_help)
    addin_upgrade_parser.add_argument("-d", "--dir", default=None, help=dir_arg_help)
    addin_upgrade_parser.set_defaults(func=addin_install)

    addin_remove_parser = addin_subparsers.add_parser("remove")
    addin_remove_parser.add_argument("-f", "--file", default=None, help=file_arg_help)
    addin_remove_parser.set_defaults(func=addin_remove)

    addin_uninstall_parser = addin_subparsers.add_parser("uninstall")
    addin_uninstall_parser.add_argument(
        "-f", "--file", default=None, help=file_arg_help
    )
    addin_uninstall_parser.set_defaults(func=addin_remove)

    addin_status_parser = addin_subparsers.add_parser("status")
    addin_status_parser.add_argument("-f", "--file", default=None, help=file_arg_help)
    addin_status_parser.set_defaults(func=addin_status)

    # Quickstart
    quickstart_parser = subparsers.add_parser(
        "quickstart",
        help='Run "xlwings quickstart myproject" to create a '
        'folder called "myproject" in the current directory '
        "with an Excel file and a Python file, ready to be "
        'used. Use the "--standalone" flag to embed all VBA '
        "code in the Excel file and make it work without the "
        "xlwings add-in.",
    )
    quickstart_parser.add_argument("project_name")
    quickstart_parser.add_argument(
        "-s", "--standalone", action="store_true", help="Include xlwings as VBA module."
    )
    quickstart_parser.add_argument(
        "-fastapi",
        "--fastapi",
        action="store_true",
        help="Create a FastAPI project suitable for a remote Python interpreter.",
    )
    quickstart_parser.add_argument(
        "-addin", "--addin", action="store_true", help="Create an add-in."
    )
    quickstart_parser.add_argument(
        "-ribbon",
        "--ribbon",
        action="store_true",
        help="Include a ribbon when creating an add-in.",
    )
    quickstart_parser.set_defaults(func=quickstart)

    # RunPython (only needed when installed with conda for Mac Excel 2016)
    if sys.platform.startswith("darwin"):
        runpython_parser = subparsers.add_parser(
            "runpython",
            help='macOS only: run "xlwings runpython install" if you '
            "want to enable the RunPython calls without installing "
            "the add-in. This will create the following file: "
            "~/Library/Application Scripts/com.microsoft.Excel/"
            "xlwings-x.x.x.applescript",
        )
        runpython_subparser = runpython_parser.add_subparsers(dest="subcommand")
        runpython_subparser.required = True

        runpython_install_parser = runpython_subparser.add_parser("install")
        runpython_install_parser.set_defaults(func=runpython_install)

    # restapi run
    restapi_parser = subparsers.add_parser(
        "restapi",
        help='Use "xlwings restapi run" to run the xlwings REST API via Flask dev '
        'server. Accepts "--host" and "--port" as optional arguments.',
    )
    restapi_subparser = restapi_parser.add_subparsers(dest="subcommand")
    restapi_subparser.required = True

    restapi_run_parser = restapi_subparser.add_parser("run")
    restapi_run_parser.add_argument(
        "-host", "--host", default="127.0.0.1", help="The interface to bind to."
    )
    restapi_run_parser.add_argument(
        "-p", "--port", default="5000", help="The port to bind to."
    )
    restapi_run_parser.set_defaults(func=restapi_run)

    # License
    license_parser = subparsers.add_parser(
        "license",
        help='xlwings PRO: Use "xlwings license update -k KEY" where '
        '"KEY" is your personal (trial) license key. This will '
        "update ~/.xlwings/xlwings.conf with the LICENSE_KEY entry. "
        'If you have a paid license, you can run "xlwings license deploy" '
        "to create a deploy key. This is not available for trial keys.",
    )
    license_subparsers = license_parser.add_subparsers(dest="subcommand")
    license_subparsers.required = True

    license_update_parser = license_subparsers.add_parser("update")
    license_update_parser.add_argument(
        "-k", "--key", help="Updates the LICENSE_KEY in ~/.xlwings/xlwings.conf."
    )
    license_update_parser.set_defaults(func=license_update)

    license_update_parser = license_subparsers.add_parser("deploy")
    license_update_parser.set_defaults(func=license_deploy)

    # Config
    config_parser = subparsers.add_parser(
        "config",
        help='Run "xlwings config create" to create the user config file '
        "(~/.xlwings/xlwings.conf) which is where the settings from "
        "the Ribbon add-in are stored. It will configure the Python "
        "interpreter that you are running this command with. To reset "
        'your configuration, run this with the "--force" flag which '
        "will overwrite your current configuration.",
    )
    config_subparsers = config_parser.add_subparsers(dest="subcommand")
    config_subparsers.required = True

    config_create_parser = config_subparsers.add_parser("create")
    config_create_parser.add_argument(
        "-f",
        "--force",
        action="store_true",
        help="Will overwrite the current config file.",
    )
    config_create_parser.set_defaults(func=config_create)

    # Embed code
    code_parser = subparsers.add_parser(
        "code",
        help='Run "xlwings code embed" to embed all Python modules of the '
        """workbook's dir in your active Excel file. Use the "--file" flag to """
        "only import a single file by providing its path. Requires "
        "xlwings PRO.",
    )
    code_subparsers = code_parser.add_subparsers(dest="subcommand")
    code_subparsers.required = True

    code_create_parser = code_subparsers.add_parser("embed")
    code_create_parser.add_argument(
        "-f",
        "--file",
        help="Optional parameter to only import a single file provided as file path.",
    )
    code_create_parser.set_defaults(func=code_embed)

    # Permission
    permission_parser = subparsers.add_parser(
        "permission",
        help='"xlwings permission cwd" prints a JSON string that can'
        " be used to permission the execution of all modules in"
        " the current working directory via GET request. "
        '"xlwings permission book" does the same for code '
        "that is embedded in the active workbook.",
    )
    permission_subparsers = permission_parser.add_subparsers(dest="subcommand")
    permission_subparsers.required = True

    permission_cwd_parser = permission_subparsers.add_parser("cwd")
    permission_cwd_parser.set_defaults(func=permission_cwd)

    permission_book_parser = permission_subparsers.add_parser("book")
    permission_book_parser.set_defaults(func=permission_book)

    # Release
    release_parser = subparsers.add_parser(
        "release",
        help='Run "xlwings release" to configure your active workbook to work with a '
        "one-click installer for easy deployment. Requires xlwings PRO.",
    )
    release_parser.set_defaults(func=release)

    # Copy
    copy_parser = subparsers.add_parser(
        "copy",
        help='Run "xlwings copy os" to copy the xlwings Office Scripts module. '
        'Run "xlwings copy gs" to copy the xlwings Google Apps Script module.',
    )
    copy_subparser = copy_parser.add_subparsers(dest="subcommand")
    copy_subparser.required = True

    copy_os_parser = copy_subparser.add_parser("os")
    copy_os_parser.set_defaults(func=copy_os)

    copy_os_parser = copy_subparser.add_parser("gs")
    copy_os_parser.set_defaults(func=copy_gs)

    # Edit VBA code
    vba_parser = subparsers.add_parser(
        "vba",
        help="""This functionality allows you to easily write VBA code in an external
        editor: run "xlwings vba edit" to update the VBA modules of the active workbook
        from their local exports everytime you hit save. If you run this the first time,
        the modules will be automatically exported from Excel. To overwrite the local
        version of the modules with the one from Excel, run "xlwings vba export".
        The "--file" flag allows you to specify a file path instead of using the active
        Workbook. Requires "Trust access to the VBA project object model" enabled.
        NOTE: Whenever you change something in the VBA editor (such as the layout of a
        form or the properties of a module), you should run "xlwings vba export" again.
        """,
    )
    vba_subparsers = vba_parser.add_subparsers(dest="subcommand")
    vba_subparsers.required = True

    vba_edit_parser = vba_subparsers.add_parser("edit")
    vba_edit_parser.add_argument(
        "-f",
        "--file",
        help="Optional parameter to select a specific workbook, otherwise it uses the "
        "active one.",
    )
    vba_edit_parser.add_argument(
        "-v",
        "--verbose",
        action="store_true",
        help="Optional parameter to print messages whenever a module has been updated "
        "successfully.",
    )

    vba_edit_parser.set_defaults(func=vba_edit)

    vba_export_parser = vba_subparsers.add_parser("export")
    vba_export_parser.add_argument(
        "-f",
        "--file",
        help="Optional parameter to select a specific file, otherwise it uses the "
        "active one.",
    )

    vba_export_parser.set_defaults(func=vba_export)

    vba_import_parser = vba_subparsers.add_parser("import")
    vba_import_parser.add_argument(
        "-f",
        "--file",
        help="Optional parameter to select a specific file, otherwise it uses the "
        "active one.",
    )

    vba_import_parser.set_defaults(func=vba_import)

    # Show help when running without commands
    if len(sys.argv) == 1:
        parser.print_help(sys.stderr)
        sys.exit(1)

    # Boilerplate
    args = parser.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
