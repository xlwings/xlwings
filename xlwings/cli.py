import os
import sys
import shutil
import argparse
from pathlib import Path

import xlwings as xw

# Directories/paths
this_dir = os.path.dirname(os.path.realpath(__file__))


def get_addin_path():
    # The call to startup_path creates the XLSTART folder if it doesn't exist yet
    if xw.apps:
        return os.path.join(xw.apps.active.startup_path, 'xlwings.xlam')
    else:
        app = xw.App(visible=False)
        startup_path = app.startup_path
        app.quit()
        return os.path.join(startup_path, 'xlwings.xlam')


def addin_install(args):
    if args is None:
        unprotected = False
    else:
        unprotected = args.unprotected
    try:
        addin_path = get_addin_path()
        addin_name = 'xlwings_unprotected.xlam' if unprotected else 'xlwings.xlam'
        shutil.copyfile(os.path.join(this_dir, 'addin', addin_name), addin_path)
        print('Successfully installed the xlwings add-in! Please restart Excel.')
        if sys.platform.startswith('darwin'):
            runpython_install(None)
        config_create(None)
    except IOError as e:
        if e.args[0] == 13:
            print('Error: Failed to install the add-in: If Excel is running, quit Excel and try again.')
        else:
            print(repr(e))
    except Exception as e:
        print(repr(e))


def addin_remove(args):
    try:
        addin_path = get_addin_path()
        os.remove(addin_path)
        print('Successfully removed the xlwings add-in!')
    except (WindowsError, PermissionError) as e:
        if e.args[0] in (13, 32):
            print('Error: Failed to remove the add-in: If Excel is running, quit Excel and try again. '
                  'You can also delete it manually from {0}'.format(addin_path))
        elif e.args[0] == 2:
            print("Error: Could not remove the xlwings add-in. The add-in doesn't seem to be installed.")
        else:
            print(repr(e))
    except Exception as e:
        print(repr(e))


def addin_status(args):
    addin_path = get_addin_path()
    if os.path.isfile(addin_path):
        print('The add-in is installed at {}'.format(addin_path))
        print('Use "xlwings addin remove" to uninstall it.')
    else:
        print('The add-in is not installed.')
        print('"xlwings addin install" will install it at: {}'.format(addin_path))


def quickstart(args):
    project_name = args.project_name
    cwd = os.getcwd()

    # Project dir
    project_path = os.path.join(cwd, project_name)
    if not os.path.exists(project_path):
        os.makedirs(project_path)
    else:
        sys.exit('Error: Directory already exists.')

    # Python file
    with open(os.path.join(project_path, project_name + '.py'), 'w') as python_module:
        python_module.write('import xlwings as xw\n\n\n')
        if sys.platform.startswith('win'):
            python_module.write('@xw.sub  # only required if you want to import it or run it via UDF Server\n')
        python_module.write('def main():\n')
        python_module.write('    wb = xw.Book.caller()\n')
        python_module.write('    sheet = wb.sheets[0]\n')
        python_module.write('    if sheet["A1"].value == "Hello xlwings!":\n')
        python_module.write('        sheet["A1"].value = "Bye xlwings!"\n')
        python_module.write('    else:\n')
        python_module.write('        sheet["A1"].value = "Hello xlwings!"\n\n\n')
        if sys.platform.startswith('win'):
            python_module.write('@xw.func\n')
            python_module.write('def hello(name):\n')
            python_module.write('    return "hello {0}".format(name)\n\n\n')
        python_module.write('if __name__ == "__main__":\n')
        python_module.write('    xw.Book("{0}.xlsm").set_mock_caller()\n'.format(project_name))
        python_module.write('    main()\n')

    # Excel file
    if not args.standalone:
        source_file = os.path.join(this_dir, 'quickstart.xlsm')
    else:
        source_file = os.path.join(this_dir, 'quickstart_standalone.xlsm')
    shutil.copyfile(source_file, os.path.join(project_path, project_name + '.xlsm'))


def runpython_install(args):
    destination_dir = os.path.expanduser("~") + '/Library/Application Scripts/com.microsoft.Excel'
    if not os.path.exists(destination_dir):
        os.makedirs(destination_dir)
    shutil.copy(os.path.join(this_dir, 'xlwings.applescript'), destination_dir)
    print('Successfully enabled RunPython!')


def restapi_run(args):
    import subprocess
    try:
        import flask
    except ImportError:
        sys.exit("To use the xlwings REST API server, you need Flask>=1.0.0 installed.")
    host = args.host
    port = args.port

    os.environ['FLASK_APP'] = 'xlwings.rest.api'
    subprocess.check_call(["flask", "run", "--host", host, "--port", port])


def license_update(args):
    """license handler for xlwings PRO"""
    key = args.key
    if not key:
        sys.exit('Please provide a license key via the -k/--key option. For example: xlwings license update -k MY_KEY')
    license_kv = '"LICENSE_KEY","{0}"\n'.format(key)
    # Update xlwings.conf
    new_config = []
    if os.path.exists(xw.USER_CONFIG_FILE):
        with open(xw.USER_CONFIG_FILE, 'r') as f:
            config = f.readlines()
        for line in config:
            # Remove existing license key and empty lines
            if line.split(',')[0] == '"LICENSE_KEY"' or line in ('\r\n', '\n'):
                pass
            else:
                new_config.append(line)
        new_config.append(license_kv)
    else:
        new_config = [license_kv]
    if not os.path.exists(os.path.dirname(xw.USER_CONFIG_FILE)):
        os.makedirs(os.path.dirname(xw.USER_CONFIG_FILE))
    with open(xw.USER_CONFIG_FILE, 'w') as f:
        f.writelines(new_config)
    print('Successfully updated license key.')


def license_deploy(args):
    from .pro import LicenseHandler
    print(LicenseHandler.create_deploy_key())


def get_conda_settings():
    conda_env = os.getenv('CONDA_DEFAULT_ENV')
    conda_exe = os.getenv('CONDA_EXE')

    if conda_env and conda_exe:
        # xlwings currently expects the path without the trailing /bin/conda or \Scripts\conda.exe
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
    if conda_path and sys.platform.startswith('win'):
        settings.append('"CONDA PATH","{}"\n'.format(conda_path))
        settings.append('"CONDA ENV","{}"\n'.format(conda_env))
    else:
        extension = 'MAC' if sys.platform.startswith('darwin') else 'WIN'
        settings.append('"INTERPRETER_{}","{}"\n'.format(extension, sys.executable))
    if os.path.exists(xw.USER_CONFIG_FILE) and not force:
        print("There is already an existing config file. Run with --force if you want to overwrite.")
    else:
        with open(xw.USER_CONFIG_FILE, 'w') as f:
            f.writelines(settings)


def code_embed(args):
    """Import all Python files of the current directory into the active Excel Book"""
    wb = xw.books.active
    screen_updating = wb.app.screen_updating
    wb.app.screen_updating = False

    if args.file:
        source_files = [Path(args.file)]
    else:
        source_files = Path('.').glob('*.py')

    for source_file in source_files:
        with open(source_file, 'r') as f:
            content = []
            for line in f.read().splitlines():
                # Handle single-quote docstrings
                line = line.replace("'''", '"""')
                # Duplicate leading single quotes so Excel interprets them properly
                # This is required even if the cell is in Text format
                content.append(["'" + line if line.startswith("'") else line])

        if source_file.name not in [sht.name for sht in wb.sheets]:
            sheet = wb.sheets.add(source_file.name, after=wb.sheets[len(wb.sheets) - 1])
        else:
            sheet = wb.sheets[source_file.name]
        sheet.cells.clear_contents()
        sheet['A1'].resize(row_size=len(content)).number_format = '@'
        sheet['A1'].value = content
        sheet['A:A'].column_width = 65

    wb.app.screen_updating = screen_updating


def main():
    print('xlwings version: ' + 'dev')
    parser = argparse.ArgumentParser()
    subparsers = parser.add_subparsers(dest='command')
    subparsers.required = True

    # Add-in
    addin_parser = subparsers.add_parser('addin', help='Run "xlwings addin install" to install the Excel add-in '
                                                       '(will be copied to the XLSTART folder). Instead of "install" you can '
                                                       'also use "update", "remove" or "status". Note that this command '
                                                       'may take a while. Use the "--unprotected" flag to install the'
                                                       'add-in without password protection.')
    addin_subparsers = addin_parser.add_subparsers(dest='subcommand')
    addin_subparsers.required = True

    addin_install_parser = addin_subparsers.add_parser('install')
    addin_install_parser.add_argument("-u", "--unprotected", action='store_true', help='Install the add-in without the password protection.')
    addin_install_parser.set_defaults(func=addin_install)

    addin_update_parser = addin_subparsers.add_parser('update')
    addin_update_parser.set_defaults(func=addin_install)

    addin_upgrade_parser = addin_subparsers.add_parser('upgrade')
    addin_upgrade_parser.set_defaults(func=addin_install)

    addin_remove_parser = addin_subparsers.add_parser('remove')
    addin_remove_parser.set_defaults(func=addin_remove)

    addin_uninstall_parser = addin_subparsers.add_parser('uninstall')
    addin_uninstall_parser.set_defaults(func=addin_remove)

    addin_status_parser = addin_subparsers.add_parser('status')
    addin_status_parser.set_defaults(func=addin_status)

    # Quickstart
    quickstart_parser = subparsers.add_parser('quickstart', help='Run "xlwings quickstart myproject" to create a '
                                                                 'folder called "myproject" in the current directory '
                                                                 'with an Excel file and a Python file, ready to be '
                                                                 'used. Use the "--standalone" flag to embed all VBA '
                                                                 'code in the Excel file and make it work without the '
                                                                 'xlwings add-in.')
    quickstart_parser.add_argument("project_name")
    quickstart_parser.add_argument("-s", "--standalone", action='store_true', help='Include xlwings as VBA module.')
    quickstart_parser.set_defaults(func=quickstart)

    # RunPython (only needed when installed with conda for Mac Excel 2016)
    if sys.platform.startswith('darwin'):
        runpython_parser = subparsers.add_parser('runpython', help='macOS only: run "xlwings runpython install" if you '
                                                                   'want to enable the RunPython calls without installing '
                                                                   'the add-in. This will create the following file: '
                                                                   '~/Library/Application Scripts/com.microsoft.Excel/xlwings.applescript')
        runpython_subparser = runpython_parser.add_subparsers(dest='subcommand')
        runpython_subparser.required = True

        runpython_install_parser = runpython_subparser.add_parser('install')
        runpython_install_parser.set_defaults(func=runpython_install)

    # restapi run
    restapi_parser = subparsers.add_parser('restapi',
                                           help='Use "xlwings restapi run" to run the xlwings REST API via Flask dev '
                                                'server. Accepts "--host" and "--port" as optional arguments.')
    restapi_subparser = restapi_parser.add_subparsers(dest='subcommand')
    restapi_subparser.required = True

    restapi_run_parser = restapi_subparser.add_parser('run')
    restapi_run_parser.add_argument("-host", "--host", default='127.0.0.1', help='The interface to bind to.')
    restapi_run_parser.add_argument("-p", "--port", default='5000', help='The port to bind to.')
    restapi_run_parser.set_defaults(func=restapi_run)

    # License
    license_parser = subparsers.add_parser('license', help='xlwings PRO: Use "xlwings license update -k KEY" where '
                                                           '"KEY" is your personal (trial) license key. This will '
                                                           'update ~/.xlwings/xlwings.conf with the LICENSE_KEY entry. '
                                                           'If you have a paid license, you can run "xlwings license deploy" '
                                                           'to create a deploy key. This is not availalbe for trial keys.')
    license_subparsers = license_parser.add_subparsers(dest='subcommand')
    license_subparsers.required = True

    license_update_parser = license_subparsers.add_parser('update')
    license_update_parser.add_argument("-k", "--key", help='Updates the LICENSE_KEY in ~/.xlwings/xlwings.conf.')
    license_update_parser.set_defaults(func=license_update)

    license_update_parser = license_subparsers.add_parser('deploy')
    license_update_parser.set_defaults(func=license_deploy)

    # Config
    config_parser = subparsers.add_parser('config', help='Run "xlwings config create" to create the user config file '
                                                         '(~/.xlwings/xlwings.conf) which is where the settings from '
                                                         'the Ribbon add-in are stored. It will configure the Python '
                                                         'interpreter that you are running this command with. To reset '
                                                         'your configuration, run this with the "--force" flag which '
                                                         'will overwrite your current configuration.')
    config_subparsers = config_parser.add_subparsers(dest='subcommand')
    config_subparsers.required = True

    config_create_parser = config_subparsers.add_parser('create')
    config_create_parser.add_argument("-f", "--force", action='store_true', help='Will overwrite the current config file.')
    config_create_parser.set_defaults(func=config_create)

    # Embed code
    code_parser = subparsers.add_parser('code', help='Run "xlwings code embed" to embed all Python modules of the '
                                                     'current dir in your active Excel file. Use the "--file" flag to '
                                                     'only import a single file by providing its path. To run embedded '
                                                     'code, you need an xlwings PRO license.')
    code_subparsers = code_parser.add_subparsers(dest='subcommand')
    code_subparsers.required = True

    code_create_parser = code_subparsers.add_parser('embed')
    code_create_parser.add_argument("-f", "--file", help='Optional parameter to only import a single file provided as file path.')
    code_create_parser.set_defaults(func=code_embed)

    # Show help when running without commands
    if len(sys.argv) == 1:
        parser.print_help(sys.stderr)
        sys.exit(1)

    # Boilerplate
    args = parser.parse_args()
    args.func(args)


if __name__ == '__main__':
    main()
