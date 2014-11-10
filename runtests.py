# Run this file on Windows only. It expects full Python installations with all dependencies in the paths below.
# On Mac, run the tests with tox.


from __future__ import print_function
import os
import inspect
from subprocess import call
from xlwings import __version__

# Python installations
py26 = r'C:\Python26'
py27 = r'C:\Python27'
py31 = r'C:\Python31'  # numpy/pandas are missing for 3.1
py32 = r'C:\Python32'
py33 = r'C:\Python33'
py34 = r'C:\Python34'


class Colors:
    # Only has an effect in the Terminal of PyCharm
    yellow = '\033[93m'
    end = '\033[0m'

this_dir = os.path.abspath(os.path.dirname(inspect.getfile(inspect.currentframe())))
setup_file = os.path.abspath(os.path.join(this_dir, 'setup.py'))

# Create distribution package
call('python {0} sdist'.format(setup_file))

# Install it, run the tests and uninstall it for each Python version
for py in [py26]:
    # Paths
    pip = os.path.abspath(os.path.join(py, 'Scripts/pip'))
    test_runner = os.path.abspath(os.path.join(py, 'Scripts/nosetests'))
    test_dir = os.path.abspath(os.path.join(py, 'Lib/site-packages/xlwings/tests'))
    xlwings_package = os.path.abspath(os.path.join(this_dir, 'dist/xlwings-{0}.zip'.format(__version__)))

    print('{0}### {1} ###{2}'.format(Colors.yellow, py, Colors.end))

    call('{0}\python -c "import sys;print(sys.version)"'.format(py))
    call('{0}\python -c "import numpy;import pandas;print(numpy.__version__);print(pandas.__version__)"'.format(py))

    # Uninstall in case there is still an existing installation
    call('{0} uninstall xlwings -y'.format(pip))

    # Install
    os.chdir(py)
    call('{0} install {1}'.format(pip, xlwings_package))

    # # Run tests
    # call('{0} {1}'.format(test_runner, test_dir))
    #
    # # Uninstall
    # call('{0} uninstall xlwings -y'.format(pip))