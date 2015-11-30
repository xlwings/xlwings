import os
import os.path as op
import sys
import shutil
import argparse
import subprocess


# Directories/paths
this_dir = os.path.dirname(os.path.realpath(__file__))
template_origin_path = os.path.join(this_dir, 'xlwings_template.xltm')

if sys.platform.startswith('win'):
    win_template_path = op.join(os.getenv('APPDATA'), 'Microsoft', 'Templates', 'xlwings_template.xltm')
else:
    # Mac 2011 and 2016 use different directories
    from appscript import k, app
    from xlwings._xlmac import hfs_to_posix_path

    mac_template_dirs = set((op.realpath(op.join(op.expanduser("~"), 'Library', 'Application Support', 'Microsoft',
                                                 'Office', 'User Templates', 'My Templates')),
                             hfs_to_posix_path(app('Microsoft Excel').properties().get(k.templates_path))))

if sys.platform.startswith('win'):
    addin_path = os.path.join(os.getenv('APPDATA'), 'Microsoft', 'Excel', 'XLSTART', 'xlwings.xlam')


def addin_install(args):
    if not sys.platform.startswith('win'):
        print('Error: This command is only available on Windows right now.')
    else:
        try:
            shutil.copyfile(os.path.join(this_dir, 'xlwings.xlam'), addin_path)
            print('Successfully installed the xlwings add-in! Please restart Excel.')
        except IOError as e:
            if e.args[0] == 13:
                print('Error: Failed to install the add-in: If Excel is running, quit Excel and try again.')
            else:
                print(str(e))
        except Exception as e:
            print(str(e))


def addin_remove(args):
    if not sys.platform.startswith('win'):
        print('Error: This command is only available on Windows right now.')
    else:
        try:
            os.remove(addin_path)
            print('Successfully removed the xlwings add-in!')
        except WindowsError as e:
            if e.args[0] == 32:
                print('Error: Failed to remove the add-in: If Excel is running, quit Excel and try again.')
            elif e.args[0] == 2:
                print("Error: Could not remove the xlwings add-in. The add-in doesn't seem to be installed.")
            else:
                print(str(e))
        except Exception as e:
            print(str(e))


def addin_status(args):
    if not sys.platform.startswith('win'):
        print('Error: This command is only available on Windows right now.')
    else:
        if os.path.isfile(addin_path):
            print('The add-in is installed at {}'.format(addin_path))
            print('Use "xlwings addin remove" to uninstall it.')
        else:
            print('The add-in is not installed.')
            print('"xlwings addin install" will install it at: {}'.format(addin_path))


def template_open(args):
    if sys.platform.startswith('win'):
        subprocess.Popen('start {0}'.format(template_origin_path), shell=True)
    else:
        subprocess.Popen('open {0}'.format(template_origin_path), shell=True)


def template_install(args):
    if sys.platform.startswith('win'):
        try:
            shutil.copyfile(template_origin_path, win_template_path)
            print('Successfully installed the xlwings template')
        except Exception as e:
            print(str(e))
    else:
        for dir in mac_template_dirs:
            try:
                if os.path.isdir(dir):
                    path = op.realpath(op.join(dir, 'xlwings_template.xltm'))
                    shutil.copyfile(template_origin_path, path)
                    print('Successfully installed the xlwings template to {}'.format(path))
            except Exception as e:
                print('Error installing template to {}. {}'.format(path, str(e)))


def template_remove(args):
    if sys.platform.startswith('win'):
        try:
            os.remove(win_template_path)
            print('Successfully removed the xlwings template!')
        except WindowsError as e:
            print("Error: Could not remove the xlwings template. The template doesn't seem to be installed.")
        except Exception as e:
            print(str(e))
    else:
        for dir in mac_template_dirs:
            try:
                if os.path.isdir(dir):
                    path = op.realpath(op.join(dir, 'xlwings_template.xltm'))
                    os.remove(path)
                    print('Successfully removed the xlwings template from {}'.format(path))
            except OSError as e:
                print("Error: Could not remove the xlwings template. "
                      "The template doesn't seem to be installed at {}.".format(path))

            except Exception as e:
                print('Error removing template from {}. {}'.format(path, str(e)))


def template_status(args):
    if sys.platform.startswith('win'):
        if os.path.isfile(win_template_path):
            print('The template is installed at: {}'.format(win_template_path))
            print ('Use "xlwings template remove" to uninstall it.')
        else:
            print('The template can be installed at {}'.format(win_template_path))
            print('Use "xlwings template install" to install it or '
                  '"xlwings template open" to open it without installing.')
    else:
        is_installed = False
        can_be_installed = False
        for dir in mac_template_dirs:
            path = op.realpath(op.join(dir, 'xlwings_template.xltm'))
            if os.path.isfile(path):
                is_installed = True
                print('The template is installed at: {}'.format(path))
            else:
                if os.path.isdir(dir):
                    can_be_installed = True
                    print('The template can be installed at: {}'.format(dir))
        if can_be_installed:
            print('Use "xlwings template install" to install it or '
                  '"xlwings template open" to open it without installing.')
        if is_installed:
            print('Use "xlwings template remove" to uninstall it from all locations.')


def main():
    parser = argparse.ArgumentParser()
    subparsers = parser.add_subparsers()

    # Add-in
    addin_parser = subparsers.add_parser('addin', help='xlwings Excel Add-in')
    addin_subparsers = addin_parser.add_subparsers()
    
    addin_install_parser = addin_subparsers.add_parser('install')
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

    # Template
    template_parser = subparsers.add_parser('template', help='xlwings Excel template')
    template_subparsers = template_parser.add_subparsers()

    template_open_parser = template_subparsers.add_parser('open')
    template_open_parser.set_defaults(func=template_open)

    template_install_parser = template_subparsers.add_parser('install')
    template_install_parser.set_defaults(func=template_install)

    template_update_parser = template_subparsers.add_parser('update')
    template_update_parser.set_defaults(func=template_install)

    template_remove_parser = template_subparsers.add_parser('remove')
    template_remove_parser.set_defaults(func=template_remove)

    template_uninstall_parser = template_subparsers.add_parser('uninstall')
    template_uninstall_parser.set_defaults(func=template_remove)

    template_status_parser = template_subparsers.add_parser('status')
    template_status_parser.set_defaults(func=template_status)

    args = parser.parse_args()
    args.func(args)

if __name__ == '__main__':
    main()
