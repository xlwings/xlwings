import os
import shutil
from subprocess import call


def main():
    # Clean the build directory
    if os.path.isdir('./build'):
        shutil.rmtree('./build')

    # Freeze it
    call('python setup_fibonacci.py build')

    # Zip it up - 7-zip provides better compression than the zipfile module
    # Make sure the 7-zip folder is on your path
    file_name = 'fibonacci_standalone'
    if os.path.isfile(file_name):
        os.remove(file_name)
    call('7z a -tzip {}.zip fibonacci.xlsm'.format(file_name))
    call('7z a -tzip {}.zip LICENSE.txt'.format(file_name))
    call('7z a -tzip {}.zip build'.format(file_name))


if __name__ == '__main__':
   main()