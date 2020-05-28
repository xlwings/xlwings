import os
import subprocess
from shlex import split

# Version string
if os.environ['GITHUB_REF'].startswith('refs/tags'):
    version_string = os.environ['GITHUB_REF'][10:]
else:
    version_string = os.environ['GITHUB_SHA'][:7]

# Installation
subprocess.check_call(split(f'python -m pip install Package/xlwings-{version_string}.tar.gz'))

# Changing the dir is required to prevent python from importing the package from the source code
os.chdir(os.path.expanduser('~'))  # e.g. /Users/runners
output = subprocess.check_output(split('python -c "import xlwings_reports;print(xlwings_reports.__version__)"'),
                                 stderr=subprocess.STDOUT).decode()
print('Version: ' + output)


