import sys
python_version = sys.version_info[:2]

if python_version < (2, 6):
    raise Exception("This version of xlrd requires Python 2.6 or above. "
                    "For older versions of Python, you can use the 0.8 series.")

av = sys.argv
if len(av) > 1 and av[1].lower() == "--egg":
    del av[1]
    from setuptools import setup
else:
    from distutils.core import setup


setup(
    name = 'xlwings',
    version = '0.1.0',
    author = 'Zoomer Analytics LLC',
    author_email = 'felix.zumstein@zoomeranalytics.com',
    url = 'http://xlwings.org/',
    packages = ['xlwings'],
    package_data={
            'xlwings': [
                'examples/*.*',
                ],
            },
    description = 'Library for developers to extract data from Microsoft Excel (tm) spreadsheet files',
    long_description = \
        "Extract data from Excel spreadsheets (.xls and .xlsx, versions 2.0 onwards) on any platform. " \
        "Pure Python (2.6, 2.7, 3.2+). Strong support for Excel dates. Unicode-aware.",
    platforms = ["Any platform -- don't need Windows"],
    license = 'BSD',
    keywords = ['xls', 'excel', 'spreadsheet', 'workbook'],
    classifiers = [
            'Development Status :: 5 - Production/Stable',
            'Intended Audience :: Developers',
            'License :: OSI Approved :: BSD License',
            'Programming Language :: Python',
            'Programming Language :: Python :: 2',
            'Programming Language :: Python :: 3',
            'Operating System :: OS Independent',
            'Topic :: Database',
            'Topic :: Office/Business',
            'Topic :: Software Development :: Libraries :: Python Modules',
            ],
    )
