import os
import re
from distutils.core import setup

# long_description: Take from README file
with open(os.path.join(os.path.dirname(__file__), 'README.rst')) as f:
    readme = f.read()

# Version Number
with open(os.path.join(os.path.dirname(__file__), 'xlwings', 'xlwings.py')) as f:
    version = re.compile(r".*__version__ = '(.*?)'", re.S).match(f.read()).group(1)

setup(
    name='xlwings',
    version=version,
    url='http://xlwings.org',
    license='BSD 3-clause',
    author='Zoomer Analytics LLC',
    author_email='felix.zumstein@zoomeranalytics.com',
    description='Make Excel fly: Interact with Excel from Python and vice versa.',
    long_description=readme,
    data_files=[('', ['README.rst'])],
    packages=['xlwings'],
    package_data={'xlwings': ['*.bas',
                              'tests/*.*']},
    platforms='Operating System :: Microsoft :: Windows',
    keywords=['xls', 'excel', 'spreadsheet', 'workbook'],
    classifiers=[
        'Development Status :: 4 - Beta',
        'Operating System :: Microsoft :: Windows',
        'Programming Language :: Python',
        'Programming Language :: Python :: 2',
        'Programming Language :: Python :: 3',
        'Topic :: Office/Business :: Financial :: Spreadsheet',
        'License :: OSI Approved :: BSD License']
)
