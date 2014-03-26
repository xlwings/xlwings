import os
from subprocess import call


def main():
    # Zip it up - 7-zip provides better compression than the zipfile module
    # Make sure the 7-zip folder is on your path
    file_name = 'database'
    if os.path.isfile('{}.zip'.format(file_name)):
        os.remove('{}.zip'.format(file_name))
    call('7z a -tzip {}.zip {}.xlsm'.format(file_name, file_name))
    call('7z a -tzip {}.zip LICENSE.txt'.format(file_name))
    call('7z a -tzip {}.zip chinook.sqlite'.format(file_name))
    call('7z a -tzip {}.zip database.py'.format(file_name))
    call('7z a -tzip {}.zip ../../xlwings/xlwings.py'.format(file_name))


if __name__ == '__main__':
   main()