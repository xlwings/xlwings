import os
import sys
import re
try:
    from setuptools import setup
except ImportError:
    from distutils.core import setup

# long_description: Take from README file
with open(os.path.join(os.path.dirname(__file__), 'README.rst')) as f:
    readme = f.read()

# Version Number
with open(os.path.join(os.path.dirname(__file__), 'xlwings', '__init__.py')) as f:
    version = re.compile(r".*__version__ = '(.*?)'", re.S).match(f.read()).group(1)

# Dependencies
if sys.platform.startswith('win'):
    install_requires = ['comtypes']
    # This places dlls next to python.exe for standard setup and in the parent folder for virtualenv
    data_files = [('', ['xlwings32.dll', 'xlwings64.dll'])]
elif sys.platform.startswith('darwin'):
    install_requires = ['psutil >= 2.0.0', 'appscript >= 1.0.1']
    data_files = [(os.path.expanduser("~") + '/Library/Application Scripts/com.microsoft.Excel', ['xlwings/xlwings.applescript'])]
else:
    on_rtd = os.environ.get('READTHEDOCS', None) == 'True'
    if on_rtd:
        data_files = []
        install_requires = []
    else:
        raise OSError("currently only Windows and OSX are supported.")

# This shouldn't be necessary anymore as we dropped official support for < 2.7 and < 3.3
if (sys.version_info[0] == 2 and sys.version_info[:2] < (2, 7)) or (sys.version_info[0] == 3 and sys.version_info[:2] < (3, 2)):
    install_requires = install_requires + ['argparse']

setup(
    name='xlwings',
    version=version,
    url='http://xlwings.org',
    license='BSD 3-clause',
    author='Zoomer Analytics LLC',
    author_email='felix.zumstein@zoomeranalytics.com',
    description='Make Excel fly: Interact with Excel from Python and vice versa.',
    long_description=readme,
    data_files=data_files,
    packages=['xlwings', 'xlwings.tests', 'xlwings.conversion'],
    package_data={'xlwings': ['xlwings.bas', 'tests/*.xlsx', 'tests/*.xlsm', 'tests/*.png', 'xlwings_template.xltm',
                              'quickstart.xlsm', 'xlwings.xlam', 'xlwings.applescript']},
    keywords=['xls', 'excel', 'spreadsheet', 'workbook', 'vba', 'macro'],
    install_requires=install_requires,
    entry_points={'console_scripts': ['xlwings=xlwings.command_line:main'],},
    classifiers=[
        'Development Status :: 4 - Beta',
        'Operating System :: Microsoft :: Windows',
        'Operating System :: MacOS :: MacOS X',
        'Programming Language :: Python',
        'Programming Language :: Python :: 2',
        'Programming Language :: Python :: 2.7',
        'Programming Language :: Python :: 3',
        'Programming Language :: Python :: 3.3',
        'Programming Language :: Python :: 3.4',
        'Programming Language :: Python :: 3.5',
        'Topic :: Office/Business :: Financial :: Spreadsheet',
        'License :: OSI Approved :: BSD License'],
    platforms=['Windows', 'Mac OS X'],
)
