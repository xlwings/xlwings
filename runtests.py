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

    print('### {} ###'.format(py))

    # Install
    os.chdir(py)
    call('{} install {}'.format(pip, xlwings_package))

    # Run tests
    call('{} {}'.format(test_runner, test_dir))

    # Uninstall
    call('{} uninstall xlwings -y'.format(pip))