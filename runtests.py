# This file expects a conda installation on your PATH. Run it to:
# - build the xlwings package
# - create conda environments that don't exist yet as specified under "envs"
# - run the nosetests against each of these environments

from __future__ import print_function
import os
import sys
import inspect
from subprocess import check_call, check_output
from xlwings import __version__


# Python versions - according yml files are under tests/conda_yml
envs = [
    ('xw27', '2.7'),
    ('xw33', '3.3'),
    ('xw34', '3.4'),
    ('xw35', '3.5'),
    ('xw35-no-opt-deps', '3.5'),
    ('xw34-pd-0.15.2', '3.4')
]

class Colors:
    yellow = '\033[93m'
    end = '\033[0m'

this_dir = os.path.abspath(os.path.dirname(inspect.getfile(inspect.currentframe())))
setup_file = os.path.abspath(os.path.join(this_dir, 'setup.py'))

# conda dirs
cmd = 'which' if sys.platform.startswith('darwin') else 'where'
conda_dir = check_output([cmd, 'conda']).decode('utf-8')
envs_dir = os.path.abspath(os.path.join(os.path.dirname(conda_dir), os.pardir, 'envs'))

# Create missing envs
for env in envs:
    if not os.path.isdir(os.path.join(envs_dir, env[0])):
        print('{0}###  Creating conda env {1} ###{2}'.format(Colors.yellow, env[0], Colors.end))
        platform = 'mac' if sys.platform.startswith('darwin') else 'win'
        check_call(['conda', 'env', 'create', '-f',
                    os.path.join(this_dir, 'xlwings', 'tests', 'conda_yml', platform, env[0] + '.yml')])

# Create distribution package
print('{0}###  Creating xlwings package ###{1}'.format(Colors.yellow, Colors.end))
check_call(['python', setup_file, 'sdist'])

# Install it and run the tests
for py in envs:
    # Paths
    if sys.platform.startswith('darwin'):
        pip = os.path.abspath(os.path.join(envs_dir, py[0], 'bin', 'pip'))
        test_runner = os.path.abspath(os.path.join(envs_dir, py[0], 'bin', 'nosetests'))
        test_dir = os.path.abspath(os.path.join(envs_dir, py[0], 'lib', 'python{}'.format(py[1]), 'site-packages', 'xlwings', 'tests'))
    else:
        pip = os.path.abspath(os.path.join(envs_dir, py[0], 'Scripts', 'pip'))
        test_runner = os.path.abspath(os.path.join(envs_dir, py[0], 'Scripts', 'nosetests'))
        test_dir = os.path.abspath(os.path.join(envs_dir, py[0], 'Lib', 'site-packages', 'xlwings', 'tests'))

    if __version__.endswith('dev'):
        __version__ = __version__[:-3] + '.dev0'

    ext = 'tar.gz' if sys.platform.startswith('darwin') else 'zip'
    xlwings_package = os.path.abspath(os.path.join(this_dir, 'dist', 'xlwings-{0}.{1}'.format(__version__, ext)))

    print('{0}### Running nosetests on {1} ###{2}'.format(Colors.yellow, py[0], Colors.end))

    # Install
    check_call([pip, 'install', 'pip', '--upgrade'])
    check_call([pip, 'uninstall', '-y', 'xlwings'])
    check_call([pip, 'install', xlwings_package, '--upgrade', '--force-reinstall', '--no-deps', '--no-cache-dir'])

    # Run tests
    check_call([test_runner, test_dir])

    # Uninstall
    check_call([pip, 'uninstall', '-y', 'xlwings'])

print('{0}### Done. ###{1}'.format(Colors.yellow, Colors.end))