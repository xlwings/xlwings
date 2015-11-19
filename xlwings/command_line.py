import os
import sys
import shutil
import argparse

this_dir = os.path.dirname(os.path.abspath(__file__))

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
            print str(e)


def addin_get_path(args):
    if not sys.platform.startswith('win'):
        print('Error: This command is only available on Windows right now.')
    else:
        if os.path.isfile(addin_path):
            print('The add-in is installed at: ' + addin_path)
        else:
            print('The add-in will be installed at: ' + addin_path)


def main():
    parser = argparse.ArgumentParser()
    subparsers = parser.add_subparsers()

    addin_parser = subparsers.add_parser('addin', help='xlwings Excel addin')
    addin_subparsers = addin_parser.add_subparsers()
    
    addin_install_parser = addin_subparsers.add_parser('install')
    addin_install_parser.set_defaults(func=addin_install)

    addin_install_parser = addin_subparsers.add_parser('update')
    addin_install_parser.set_defaults(func=addin_install)

    addin_install_parser = addin_subparsers.add_parser('upgrade')
    addin_install_parser.set_defaults(func=addin_install)

    addin_remove_parser = addin_subparsers.add_parser('remove')
    addin_remove_parser.set_defaults(func=addin_remove)    

    addin_remove_parser = addin_subparsers.add_parser('uninstall')
    addin_remove_parser.set_defaults(func=addin_remove)

    addin_remove_parser = addin_subparsers.add_parser('path')
    addin_remove_parser.set_defaults(func=addin_get_path)

    args = parser.parse_args()
    args.func(args)

if __name__ == '__main__':
    main()
