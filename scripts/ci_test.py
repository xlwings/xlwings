import os
import subprocess
from shlex import split
import glob

os.chdir('Package')

# Version numbers get sometimes normalized from setuptools, so just check what package is in the directory
for package in glob.glob('*.tar.gz'):
    # Installation
    subprocess.check_call(split(f'python -m pip install {package}'))

# Changing the dir is required to prevent python from importing the package from the source code
os.chdir(os.path.expanduser('~'))  # e.g. /Users/runners
output = subprocess.check_output(split('python -c "import xlwings;print(xlwings.__version__)"'),
                                 stderr=subprocess.STDOUT).decode()
print('Version: ' + output)


