import argparse
import json
import os
import shutil
import subprocess
import sys
import time
import uuid
from keyword import iskeyword
from pathlib import Path

import xlwings as xw

# Directories/paths
this_dir = Path(__file__).resolve().parent


def auth_aad(args):
    _auth_aad(
        tenant_id=args.tenant_id,
        client_id=args.client_id,
        port=args.port,
        scopes=args.scopes,
        username=args.username,
        reset=args.reset,
    )


def _auth_aad(
    client_id=None, tenant_id=None, username=None, port=None, scopes=None, reset=False
):
    from xlwings.utils import read_user_config

    try:
        import msal
    except ImportError:
        sys.exit("Couldn't find the 'msal' package. Install it via `pip install msal`.")

    cache_dir = Path(xw.USER_CONFIG_FILE).parent
    cache_file = cache_dir / "aad.json"

    if reset:
        if cache_file.exists():
            cache_file.unlink()
        update_user_config("AZUREAD_ACCESS_TOKEN", None, action="delete")
        update_user_config("AZUREAD_ACCESS_TOKEN_EXPIRES_ON", None, action="delete")

    user_config = read_user_config()
    if tenant_id is None:
        tenant_id = user_config["azuread_tenant_id"]
    if client_id is None:
        client_id = user_config["azuread_client_id"]
    if scopes is None:
        scopes = user_config["azuread_scopes"]
    if port is None:
        port = user_config.get("azuread_port")
    if username is None:
        username = user_config.get("azuread_username")
    # Scopes can only be from one application!
    if scopes is None:
        scopes = [""]
    elif isinstance(scopes, str):
        scopes = [scope.strip() for scope in scopes.split(",")]
    else:
        sys.exit("Please provide scopes as a single string with commas.")

    # https://learn.microsoft.com/en-us/azure/active-directory/develop/msal-client-application-configuration#authority
    authority = f"https://login.microsoftonline.com/{tenant_id}"

    # Cache
    token_cache = msal.SerializableTokenCache()

    cache_file.parent.mkdir(exist_ok=True)
    if cache_file.exists():
        token_cache.deserialize(cache_file.read_text())

    app = msal.PublicClientApplication(
        client_id=client_id,
        authority=authority,
        token_cache=token_cache,
    )

    result = None
    # Username selects the account if multiple accounts are logged in
    # This requires scopes to be set though, but scopes=[""] seem to do the trick.
    accounts = app.get_accounts(username=username if username else None)
    if accounts:
        account = accounts[0]
        result = app.acquire_token_silent(scopes=scopes, account=account)
    if not result:
        result = app.acquire_token_interactive(
            scopes=scopes, timeout=60, port=int(port) if port else None
        )

    if "access_token" not in result:
        sys.exit(result.get("error_description"))

    update_user_config(f"AZUREAD_ACCESS_TOKEN_{client_id}", result["access_token"])
    update_user_config(
        f"AZUREAD_ACCESS_TOKEN_EXPIRES_ON_{client_id}",
        int(time.time()) + result["expires_in"],
    )

    if token_cache.has_state_changed:
        cache_file.write_text(token_cache.serialize())


def exit_unsupported_platform():
    if not sys.platform.startswith("win"):
        sys.exit("Error: This command is currently only supported on Windows.")


def get_addin_dir(global_location=False):
    # The call to startup_path creates the XLSTART folder if it doesn't exist yet
    # The global XLSTART folder seems to be always existing
    if xw.apps:
        if global_location:
            addin_dir = Path(xw.apps.active.path) / "XLSTART"
            addin_dir.mkdir(exist_ok=True)
            return addin_dir
        else:
            return xw.apps.active.startup_path
    else:
        with xw.App(visible=False) as app:
            if global_location:
                addin_dir = Path(app.path) / "XLSTART"
                addin_dir.mkdir(exist_ok=True)
                return addin_dir
            else:
                return app.startup_path


def _handle_addin_glob_arg(args):
    if args.glob:
        if not sys.platform.startswith("win"):
            sys.exit("Error: The '--glob' option is only supported on Windows.")
        return True
    else:
        return False


def addin_install(args):
    global_install = _handle_addin_glob_arg(args)
    if args.dir:
        for addin_source_path in Path(args.dir).resolve().glob("[!~$]*.xl*"):
            _addin_install(str(addin_source_path), global_install)
    elif args.file:
        addin_source_path = os.path.abspath(args.file)
        _addin_install(addin_source_path, global_install)
    else:
        addin_source_path = os.path.join(this_dir, "addin", "xlwings.xlam")
        _addin_install(addin_source_path, global_install)


def _addin_install(addin_source_path, global_location=False):
    addin_name = os.path.basename(addin_source_path)
    addin_target_path = os.path.join(get_addin_dir(global_location), addin_name)

    # Close any open add-ins
    if xw.apps:
        for app in xw.apps:
            try:
                app.books[addin_name].close()
            except KeyError:
                pass
    try:
        # Install the add-in
        shutil.copyfile(addin_source_path, addin_target_path)
        # Load the add-in
        if xw.apps:
            for app in xw.apps:
                if sys.platform.startswith("win"):
                    try:
                        app.books.open(Path(app.startup_path) / addin_name)
                        print("Successfully installed the xlwings add-in!")
                    except:  # noqa: E722
                        print(
                            "Successfully installed the xlwings add-in! "
                            "Please restart Excel."
                        )
                    try:
                        app.activate(steal_focus=True)
                    except:  # noqa: E722
                        pass
                else:
                    # macOS asks to explicitly enable macros when opened directly
                    # which isn't a good UX
                    print(
                        "Successfully installed the xlwings add-in! "
                        "Please restart Excel (quit via Cmd-Q, then start Excel again)."
                    )
        else:
            print("Successfully installed the xlwings add-in! ")
        if sys.platform.startswith("darwin"):
            runpython_install(None)
        if addin_name == "xlwings.xlam":
            config_create(None)
    except IOError as e:
        if e.args[0] == 13:
            print(
                "Error: Failed to install the add-in: If Excel is running, "
                "quit Excel and try again. If you are using the '--glob' option "
                "make sure to run this command from an Elevated Command Prompt."
            )
        else:
            print(repr(e))
    except Exception as e:
        print(repr(e))


def addin_remove(args):
    global_install = _handle_addin_glob_arg(args)
    if args.dir:
        for addin_source_path in Path(args.dir).resolve().glob("[!~$]*.xl*"):
            _addin_remove(addin_source_path, global_install)
    elif args.file:
        _addin_remove(args.file, global_install)
    else:
        _addin_remove("xlwings.xlam", global_install)


def _addin_remove(addin_name, global_install):
    addin_name = os.path.basename(addin_name)
    addin_path = os.path.join(get_addin_dir(global_install), addin_name)
    try:
        if xw.apps:
            for app in xw.apps:
                try:
                    app.books[addin_name].close()
                except KeyError:
                    pass
        os.remove(addin_path)
        print("Successfully removed the add-in!")
    except (WindowsError, PermissionError) as e:
        if e.args[0] in (13, 32):
            print(
                "Error: Failed to remove the add-in: If Excel is running, "
                "quit Excel and try again. If you use the '--glob' option, make "
                "sure to run this command from an Elevated Command Prompt!"
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
    global_install = _handle_addin_glob_arg(args)
    if args.file:
        addin_name = os.path.basename(args.file)
    else:
        addin_name = "xlwings.xlam"
    addin_path = os.path.join(get_addin_dir(global_install), addin_name)
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
            "manually after running this command, if you also adjust your RunPython "
            "VBA function accordingly."
        )
    cwd = os.getcwd()

    # Project dir
    project_path = os.path.join(cwd, project_name)
    if args.fastapi:
        # Raises an error on its own if the dir already exists
        shutil.copytree(
            Path(this_dir) / "quickstart_fastapi",
            Path(cwd) / project_name,
            ignore=shutil.ignore_patterns("__pycache__"),
        )
    else:
        if not os.path.exists(project_path):
            os.makedirs(project_path)
        else:
            sys.exit("Error: Directory already exists.")

    # Python file
    if not args.fastapi:
        with open(os.path.join(project_path, project_name + ".py"), "w") as f:
            f.write("import xlwings as xw\n\n\n")
            f.write("def main():\n")
            f.write("    wb = xw.Book.caller()\n")
            f.write("    sheet = wb.sheets[0]\n")
            f.write('    if sheet["A1"].value == "Hello xlwings!":\n')
            f.write('        sheet["A1"].value = "Bye xlwings!"\n')
            f.write("    else:\n")
            f.write('        sheet["A1"].value = "Hello xlwings!"\n\n\n')
            if sys.platform.startswith("win"):
                f.write("@xw.func\n")
                f.write("def hello(name):\n")
                f.write('    return f"Hello {name}!"\n\n\n')
            f.write('if __name__ == "__main__":\n')
            f.write('    xw.Book("{0}.xlsm").set_mock_caller()\n'.format(project_name))
            f.write("    main()\n")

    # Excel file
    if args.standalone:
        source_file = os.path.join(this_dir, "quickstart_standalone.xlsm")
    elif args.addin and args.ribbon:
        source_file = os.path.join(this_dir, "quickstart_addin_ribbon.xlam")
    elif args.addin:
        source_file = os.path.join(this_dir, "quickstart_addin.xlam")
    else:
        source_file = os.path.join(this_dir, "quickstart.xlsm")

    target_file = os.path.join(
        project_path, project_name + os.path.splitext(source_file)[1]
    )
    shutil.copyfile(
        source_file,
        target_file,
    )

    if args.standalone and args.fastapi:
        book = xw.Book(target_file)
        import_remote_modules(book)
        book.save()


def runpython_install(args):
    destination_dir = (
        os.path.expanduser("~") + "/Library/Application Scripts/com.microsoft.Excel"
    )
    if not os.path.exists(destination_dir):
        os.makedirs(destination_dir)
    shutil.copy(
        os.path.join(this_dir, f"xlwings-{xw.__version__}.applescript"), destination_dir
    )
    if args:
        # Don't print when called as part of "xlwings addin install"
        print("Successfully enabled RunPython!")


def restapi_run(args):
    import subprocess

    try:
        import flask  # noqa: F401
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
    update_user_config("LICENSE_KEY", key)
    print("Successfully updated license key.")


def update_user_config(key, value=None, action="add"):
    # action: 'add' or 'remove'
    new_config = []
    if os.path.exists(xw.USER_CONFIG_FILE):
        with open(xw.USER_CONFIG_FILE, "r") as f:
            config = f.readlines()
        for line in config:
            # Remove existing key and empty lines
            if line.split(",")[0] == f'"{key}"' or line in ("\r\n", "\n"):
                pass
            else:
                new_config.append(line)
        if action == "add":
            new_config.append(f'"{key}","{value}"\n')
    else:
        if action == "add":
            new_config = [f'"{key}","{value}"\n']
        else:
            return
    if not os.path.exists(os.path.dirname(xw.USER_CONFIG_FILE)):
        os.makedirs(os.path.dirname(xw.USER_CONFIG_FILE))
    with open(xw.USER_CONFIG_FILE, "w") as f:
        f.writelines(new_config)


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
    if not os.path.exists(xw.USER_CONFIG_FILE) or force:
        with open(xw.USER_CONFIG_FILE, "w") as f:
            f.writelines(settings)


def code_embed(args):
    """Import a specific file or all Python files of the Excel books' directory
    into the active Excel Book
    """
    wb = xw.books.active
    single_file = False
    if args and args.file:
        source_files = [Path(args.file)]
        single_file = True
    else:
        source_files = list(Path(wb.fullname).resolve().parent.rglob("*.py"))
    if not source_files:
        print("WARNING: Couldn't find any Python files in the workbook's directory!")

    # Delete existing source code sheets
    # A bug prevents deleting sheets from the collection directly (#1400)
    with wb.app.properties(screen_updating=False):
        if not single_file:
            for sheetname in [sheet.name for sheet in wb.sheets]:
                if wb.sheets[sheetname].name.endswith(".py"):
                    wb.sheets[sheetname].delete()

        # Import source code
        sheetname_to_path = {}
        for source_file in source_files:
            if not single_file:
                sheetname = uuid.uuid4().hex[:28] + ".py"
                sheetname_to_path[sheetname] = str(
                    source_file.relative_to(Path(wb.fullname).parent)
                )
            with open(source_file, "r", encoding="utf-8") as f:
                content = []
                for line in f.read().splitlines():
                    # Handle single-quote docstrings
                    line = line.replace("'''", '"""')
                    # Duplicate leading single quotes so Excel interprets them properly
                    # This is required even if the cell is in Text format
                    content.append(["'" + line if line.startswith("'") else line])
            if single_file and source_file.name not in wb.sheet_names:
                sheet = wb.sheets.add(
                    source_file.name, after=wb.sheets[len(wb.sheets) - 1]
                )
            elif single_file:
                sheet = wb.sheets[source_file.name]
            else:
                sheet = wb.sheets.add(sheetname, after=wb.sheets[len(wb.sheets) - 1])

            sheet["A1"].resize(
                row_size=len(content) if content else 1
            ).number_format = "@"
            sheet["A1"].value = content
            sheet["A:A"].column_width = 65

        # Update config with the sheetname_to_path mapping
        if "xlwings.conf" in wb.sheet_names and not single_file:
            config = xw.utils.read_config_sheet(wb)
            sheetname_to_path_str = json.dumps(sheetname_to_path)
            if len(sheetname_to_path_str) > 32_767:
                sys.exit("ERROR: The package structure is too complex to embed.")
            config["RELEASE_EMBED_CODE_MAP"] = sheetname_to_path_str
            wb.sheets["xlwings.conf"]["A1"].value = config
        elif not single_file:
            config_sheet = wb.sheets.add(
                "xlwings.conf", after=wb.sheets[len(wb.sheets) - 1]
            )
            config_sheet["A1"].value = {
                "RELEASE_EMBED_CODE_MAP": json.dumps(sheetname_to_path)
            }


def copy_os(args):
    copy_code(Path(this_dir) / "js" / "xlwings.ts")


def copy_gs(args):
    copy_code(Path(this_dir) / "js" / "xlwings.js")


def copy_vba(args):
    if args.addin:
        copy_code(Path(this_dir) / "xlwings_custom_addin.bas")
    else:
        copy_code(Path(this_dir) / "xlwings.bas")


def copy_customaddin(args):
    copy_code(Path(this_dir) / "xlwings_custom_addin.bas")


def copy_code(fpath):
    try:
        from pandas.io import clipboard
    except ImportError:
        try:
            import pyperclip as clipboard
        except ImportError:
            sys.exit(
                'Please install either "pandas" or "pyperclip" to use the copy command.'
            )

    with open(fpath, "r", encoding="utf-8") as f:
        if "bas" in str(fpath):
            text = (
                f.read()
                .replace('Attribute VB_Name = "xlwings"\n', "")
                .replace('Attribute VB_Name = "xlwings"\r\n', "")
            )
        else:
            text = f.read()
        clipboard.copy(text)
        print("Successfully copied to clipboard.")


def import_remote_modules(book):
    for vba_module in [
        "IWebAuthenticator.cls",
        "WebClient.cls",
        "WebRequest.cls",
        "WebResponse.cls",
        "WebHelpers.bas",
    ]:
        book.api.VBProject.VBComponents.Import(this_dir / "addin" / vba_module)


def release(args):
    from xlwings.pro import LicenseHandler
    from xlwings.utils import query_yes_no, read_user_config

    if sys.platform.startswith("darwin"):
        sys.exit(
            "This command is currently only supported on Windows. "
            "However, a released workbook will work on macOS, too."
        )

    if xw.apps:
        book = xw.apps.active.books.active
    else:
        sys.exit("Please open your Excel file first.")

    # Deploy Key
    try:
        deploy_key = LicenseHandler.create_deploy_key()
    except xw.LicenseError:
        # Can't create deploy key with trial keys, so use trial key directly
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
        use_remote = query_yes_no("Support remote interpreter?", "no")
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
                "RELEASE_REMOTE_INTERPRETER": use_remote,
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

        for vba_module in [
            "xlwings",
            "Dictionary",
            "IWebAuthenticator",
            "WebClient",
            "WebRequest",
            "WebResponse",
            "WebHelpers",
        ]:
            if vba_module in [i.Name for i in book.api.VBProject.VBComponents]:
                book.api.VBProject.VBComponents.Remove(
                    book.api.VBProject.VBComponents(vba_module)
                )

        # Import VBA modules/classes
        book.api.VBProject.VBComponents.Import(this_dir / "xlwings.bas")
        book.api.VBProject.VBComponents.Import(this_dir / "addin" / "Dictionary.cls")
        if config["RELEASE_REMOTE_INTERPRETER"]:
            import_remote_modules(book)

    # Embed code
    if config.get("RELEASE_EMBED_CODE"):
        print("* Embed Python code")
        code_embed(None)

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
            [
                interpreter_path,
                "-c",
                "import warnings;warnings.filterwarnings('ignore');"
                "import xlwings;print(xlwings.__version__)",
            ],
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
            Path(".").resolve()
            / f"{vb_component.Name}.{type_to_ext[vb_component.Type]}"
        )
        path_to_type[str(file_path)] = vb_component.Type
        if (
            vb_component.Type == 100 and vb_component.CodeModule.CountOfLines > 0
        ) or vb_component.Type != 100:
            # Prevents cluttering everything with empty files if you have lots of sheets
            if overwrite or not file_path.exists():
                vb_component.Export(str(file_path))
                if vb_component.Type == 100:
                    # Remove the meta info so it can be distinguished from regular
                    # classes when running "xlwings vba import"
                    with open(file_path, "r", encoding="utf-8") as f:
                        exported_code = f.readlines()
                    with open(file_path, "w", encoding="utf-8") as f:
                        f.writelines(exported_code[9:])
    return path_to_type


def vba_get_book(args):
    import textwrap

    from xlwings.utils import query_yes_no

    if args and args.file:
        book = xw.Book(args.file)
    else:
        if not xw.apps:
            sys.exit(
                "Your workbook must be open or you have to supply the --file argument."
            )
        else:
            book = xw.books.active

    tf = query_yes_no(
        textwrap.dedent(
            f"""
    This will affect the following workbook/folder:

    * {book.name}
    * {Path(".").resolve()}

    Proceed?"""
        )
    )

    if not tf:
        sys.exit()
    return book


def vba_import(args):
    exit_unsupported_platform()
    import pywintypes

    book = vba_get_book(args)

    for path in Path(".").resolve().glob("*"):
        if path.suffix == ".bas":
            try:
                vb_component = book.api.VBProject.VBComponents(path.stem)
                book.api.VBProject.VBComponents.Remove(vb_component)
            except pywintypes.com_error:
                pass
            book.api.VBProject.VBComponents.Import(path)
        elif path.suffix in (".cls", ".frm"):
            with open(path, "r", encoding="utf-8") as f:
                vba_code = f.readlines()
            if vba_code:
                if vba_code[0].startswith("VERSION "):
                    # For frm, this also imports frx, unlike in editing mode
                    try:
                        vb_component = book.api.VBProject.VBComponents(path.stem)
                        book.api.VBProject.VBComponents.Remove(vb_component)
                    except pywintypes.com_error:
                        pass
                    book.api.VBProject.VBComponents.Import(path)
                else:
                    vb_component = book.api.VBProject.VBComponents(path.stem)
                    line_count = vb_component.CodeModule.CountOfLines
                    if line_count > 0:
                        vb_component.CodeModule.DeleteLines(1, line_count)
                    vb_component.CodeModule.AddFromString("".join(vba_code))
    book.save()


def vba_export(args):
    exit_unsupported_platform()
    book = vba_get_book(args)
    export_vba_modules(book, overwrite=True)
    print(f"Successfully exported the VBA modules from {book.name}!")


def vba_edit(args):
    exit_unsupported_platform()
    import pywintypes

    try:
        from watchgod import Change, RegExpWatcher, watch
    except ImportError:
        sys.exit(
            "Please install watchgod to use this functionality: pip install watchgod"
        )

    book = vba_get_book(args)

    path_to_type = export_vba_modules(book, overwrite=False)
    mode = "verbose" if args.verbose else "silent"

    print("NOTE: Deleting a VBA module here will also delete it in the VBA editor!")
    print(f"Watching for changes in {book.name} ({mode} mode)...(Hit Ctrl-C to stop)")

    for changes in watch(
        Path(".").resolve(),
        watcher_cls=RegExpWatcher,
        watcher_kwargs=dict(re_files=r"^.*(\.cls|\.frm|\.bas)$"),
        normal_sleep=400,
    ):
        for change_type, path in changes:
            module_name = os.path.splitext(os.path.basename(path))[0]
            module_type = path_to_type[path]
            vb_component = book.api.VBProject.VBComponents(module_name)
            if change_type == Change.modified:
                with open(path, "r", encoding="utf-8") as f:
                    vba_code = f.readlines()
                line_count = vb_component.CodeModule.CountOfLines
                if line_count > 0:
                    vb_component.CodeModule.DeleteLines(1, line_count)
                # ThisWorkbook/Sheet, bas, cls, frm
                type_to_firstline = {100: 0, 1: 1, 2: 9, 3: 15}
                try:
                    vb_component.CodeModule.AddFromString(
                        "".join(vba_code[type_to_firstline[module_type] :])
                    )
                except pywintypes.com_error:
                    print(
                        f"ERROR: Couldn't update module {module_name}. "
                        f"Please update changes manually."
                    )
                if args.verbose:
                    print(f"INFO: Updated module {module_name}.")
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


def py_edit(args):
    try:
        from watchgod import RegExpWatcher, watch
    except ImportError:
        sys.exit(
            "Please install watchgod to use this functionality: pip install watchgod"
        )
    book = xw.books.active
    selection = book.selection
    source_path = (
        selection.get_address(include_sheetname=True, external=True)
        .replace("!", "")
        .replace(" ", "_")
        .replace("'", "")
        .replace("[", "")
        .replace("]", "_")
        + ".py"
    )

    Path(source_path).write_text(selection.formula.strip()[5:-4].replace('""', '"'))
    print(f"Open {Path(source_path).resolve()} to edit!")
    print("Syncing changes... (Hit Ctrl-C to stop)")
    for changes in watch(
        Path(".").resolve(),
        watcher_cls=RegExpWatcher,
        watcher_kwargs=dict(re_files=r"^.*(\.py)$"),
        normal_sleep=400,
    ):
        source_code = Path(source_path).read_text()
        # 1 = Object, 0 = Value
        # Note that the initial beta version only supports 1
        selection.value = '=PY("{0}",1)'.format(source_code.replace('"', '""'))


def main():
    print("xlwings version: " + xw.__version__)
    parser = argparse.ArgumentParser()
    subparsers = parser.add_subparsers(dest="command")
    subparsers.required = True

    # Add-in
    addin_parser = subparsers.add_parser(
        "addin",
        help='Run "xlwings addin install" to install the Excel add-in '
        '(will be copied to the user\'s XLSTART folder). Instead of "install" you can '
        'also use "update", "remove" or "status". Note that this command '
        "may take a while. You can install your custom add-in "
        "by providing the name or path via the --file/-f flag, "
        'e.g. "xlwings add-in install -f custom.xlam or copy all Excel '
        "files in a directory to the XLSTART folder by providing the path "
        'via the --dir flag." To install the add-in for every user globally, use the '
        " --glob/-g flag and run this command from an Elevated Command Prompt.",
    )
    addin_subparsers = addin_parser.add_subparsers(dest="subcommand")
    addin_subparsers.required = True

    file_arg_help = "The name or path of a custom add-in."
    dir_arg_help = (
        "The path of a directory whose Excel files you want to copy to or remove from "
        "XLSTART."
    )
    glob_arg_help = "Install the add-in for all users."

    addin_install_parser = addin_subparsers.add_parser("install")

    addin_install_parser.add_argument("-f", "--file", default=None, help=file_arg_help)
    addin_install_parser.add_argument("-d", "--dir", default=None, help=dir_arg_help)
    addin_install_parser.add_argument(
        "-g", "--glob", action="store_true", help=glob_arg_help
    )
    addin_install_parser.set_defaults(func=addin_install)

    addin_update_parser = addin_subparsers.add_parser("update")
    addin_update_parser.add_argument("-f", "--file", default=None, help=file_arg_help)
    addin_update_parser.add_argument("-d", "--dir", default=None, help=dir_arg_help)
    addin_update_parser.add_argument(
        "-g", "--glob", action="store_true", help=glob_arg_help
    )
    addin_update_parser.set_defaults(func=addin_install)

    addin_upgrade_parser = addin_subparsers.add_parser("upgrade")
    addin_upgrade_parser.add_argument("-f", "--file", default=None, help=file_arg_help)
    addin_upgrade_parser.add_argument("-d", "--dir", default=None, help=dir_arg_help)
    addin_upgrade_parser.add_argument(
        "-g", "--glob", action="store_true", help=glob_arg_help
    )
    addin_upgrade_parser.set_defaults(func=addin_install)

    addin_remove_parser = addin_subparsers.add_parser("remove")
    addin_remove_parser.add_argument("-f", "--file", default=None, help=file_arg_help)
    addin_remove_parser.add_argument("-d", "--dir", default=None, help=dir_arg_help)
    addin_remove_parser.add_argument(
        "-g", "--glob", action="store_true", help=glob_arg_help
    )
    addin_remove_parser.set_defaults(func=addin_remove)

    addin_uninstall_parser = addin_subparsers.add_parser("uninstall")
    addin_uninstall_parser.add_argument(
        "-f", "--file", default=None, help=file_arg_help
    )
    addin_uninstall_parser.add_argument("-d", "--dir", default=None, help=dir_arg_help)
    addin_uninstall_parser.add_argument(
        "-g", "--glob", action="store_true", help=glob_arg_help
    )
    addin_uninstall_parser.set_defaults(func=addin_remove)

    addin_status_parser = addin_subparsers.add_parser("status")
    addin_status_parser.add_argument("-f", "--file", default=None, help=file_arg_help)
    addin_status_parser.add_argument(
        "-g", "--glob", action="store_true", help=glob_arg_help
    )
    addin_status_parser.set_defaults(func=addin_status)

    # Quickstart
    quickstart_parser = subparsers.add_parser(
        "quickstart",
        help='Run "xlwings quickstart myproject" to create a '
        'folder called "myproject" in the current directory '
        "with an Excel file and a Python file, ready to be "
        'used. Use the "--standalone" flag to embed all VBA '
        "code in the Excel file and make it work without the "
        "xlwings add-in. "
        'Use "--fastapi" for creating a project that uses a remote '
        "Python interpreter. "
        'Use "--addin --ribbon" to create a template for a custom ribbon addin. Leave '
        'away the "--ribbon" if you don\'t want a ribbon tab. ',
    )
    quickstart_parser.add_argument("project_name")
    quickstart_parser.add_argument(
        "-s", "--standalone", action="store_true", help="Include xlwings as VBA module."
    )
    quickstart_parser.add_argument(
        "-r",
        "--remote",
        action="store_true",
        help="Support a remote Python interpreter.",
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

    # RunPython (macOS only)
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
        help='Use "xlwings restapi run" to run the xlwings REST API via Flask '
        'development server. Accepts "--host" and "--port" as optional arguments.',
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
        """workbook's dir in your active Excel file. Use the "--file/-f" flag to """
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
        'Run "xlwings copy gs" to copy the xlwings Google Apps Script module. '
        'Run "xlwings copy vba" to copy the standalone xlwings VBA module. '
        'Run "xlwings copy vba --addin" to copy the xlwings VBA module for custom '
        "add-ins.",
    )
    copy_subparser = copy_parser.add_subparsers(dest="subcommand")
    copy_subparser.required = True

    copy_os_parser = copy_subparser.add_parser("os")
    copy_os_parser.set_defaults(func=copy_os)

    copy_os_parser = copy_subparser.add_parser("gs")
    copy_os_parser.set_defaults(func=copy_gs)

    copy_vba_parser = copy_subparser.add_parser("vba")
    copy_vba_parser.add_argument(
        "-a", "--addin", action="store_true", help="VBA for custom add-ins"
    )
    copy_vba_parser.set_defaults(func=copy_vba)

    # Azure AD authentication (MSAL)
    auth_parser = subparsers.add_parser(
        "auth",
        help='Microsoft Azure AD: "xlwings auth azuread", see '
        "https://docs.xlwings.org/en/stable/server_authentication.html",
    )

    aad_subparser = auth_parser.add_subparsers(dest="subcommand")
    aad_subparser.required = True

    auth_aad_parser = aad_subparser.add_parser("azuread")
    auth_aad_parser.set_defaults(func=auth_aad)
    auth_aad_parser.add_argument(
        "-tid",
        "--tenant_id",
        help="Tenant ID",
    )
    auth_aad_parser.add_argument(
        "-cid",
        "--client_id",
        help="CLIENT ID",
    )
    auth_aad_parser.add_argument(
        "-p",
        "--port",
        help="Port",
    )
    auth_aad_parser.add_argument(
        "-s",
        "--scopes",
        help="Scopes",
    )
    auth_aad_parser.add_argument(
        "-u",
        "--username",
        help="Username",
    )
    auth_aad_parser.add_argument(
        "-r", "--reset", action="store_true", help="Clear local cache."
    )
    # Edit VBA code
    vba_parser = subparsers.add_parser(
        "vba",
        help="""This functionality allows you to easily write VBA code in an external
        editor: run "xlwings vba edit" to update the VBA modules of the active workbook
        from their local exports everytime you hit save. If you run this the first time,
        the modules will be exported from Excel into your current working directory.
        To overwrite the local version of the modules with those from Excel,
        run "xlwings vba export". To overwrite the VBA modules in Excel with their local
        versions, run "xlwings vba import".
        The "--file/-f" flag allows you to specify a file path instead of using the
        active Workbook. Requires "Trust access to the VBA project object model"
        enabled.
        NOTE: Whenever you change something in the VBA editor (such as the layout of a
        form or the properties of a module), you have to run "xlwings vba export".
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

    # Edit =PY cells
    py_parser = subparsers.add_parser(
        "py",
        help="""This functionality allows you to easily write Python code for 
        Microsoft's Python in Excel cells (=PY) via a local editor: run "xlwings py edit" to
        export the code of the selected cell into a local file. Whenever you save the
        file, the code will be synced back to the cell.
        """,
    )
    py_subparsers = py_parser.add_subparsers(dest="subcommand")
    py_subparsers.required = True

    py_edit_parser = py_subparsers.add_parser("edit")
    py_edit_parser.set_defaults(func=py_edit)

    # Show help when running without commands
    if len(sys.argv) == 1:
        parser.print_help(sys.stderr)
        sys.exit(1)

    # Boilerplate
    args, _ = parser.parse_known_args()
    args.func(args)


if __name__ == "__main__":
    main()
