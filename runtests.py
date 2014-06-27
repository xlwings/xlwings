# Since I couldn't make tox or conda environments work for all Python versions including pywin32/numpy/pandas,
# I had to fall back to clean full installations for now.
from __future__ import print_function
import os
from subprocess import call
from xlwings import __version__

# Python installations
py26 = r'C:\Python26'
py27 = r'C:\Python27'
py31 = r'C:\Python31'  # numpy/pandas are missing for 3.1
py32 = r'C:\Python32'
py33 = r'C:\Python33'
py34 = r'C:\Python34'


class colors:
    yellow = '\033[93m'
    end = '\033[0m'

this_path = os.path.dirname(__file__)
setup_file = os.path.abspath(os.path.join(this_path, 'setup.py'))

# Create distribution package
call('python {} sdist'.format(setup_file))

# Install it, run the tests and uninstall it for each Python version
for py in [py26, py27, py31, py32, py33, py34]:
    # Paths
    pip = os.path.abspath(os.path.join(py, 'Scripts/pip'))
    test_runner = os.path.abspath(os.path.join(py, 'Scripts/nosetests'))
    test_dir = os.path.abspath(os.path.join(py, 'Lib/site-packages/xlwings/tests'))
    xlwings_package = os.path.abspath(os.path.join(this_path, 'dist/xlwings-{}.zip'.format(__version__)))

    print('{}### {} ###{}'.format(colors.yellow, py, colors.end))

    call('{}\python -c "import sys;print(sys.version)"'.format(py))
    call('{}\python -c "import numpy;import pandas;print(numpy.__version__);print(pandas.__version__)"'.format(py))

    # Uninstall in case there is still an existing installation
    call('{} uninstall xlwings -y'.format(pip))

    # Install
    os.chdir(py)
    call('{} install {}'.format(pip, xlwings_package))

    # Run tests
    call('{} {}'.format(test_runner, test_dir))

    # Uninstall
    call('{} uninstall xlwings -y'.format(pip))