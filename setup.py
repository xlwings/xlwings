from distutils.core import setup

setup(
    name='xlwings',
    version='0.1.0',
    url='http://xlwings.org',
    license='BSD 3-clause',
    author='Zoomer Analytics LLC',
    author_email='felix.zumstein@zoomeranalytics.com',
    description='Make Excel fly: Interact with Excel from Python and vice versa.',
    long_description='TODO',
    packages=['xlwings'],
    package_data={'xlwings': ['*.bas',
                              'examples/*.*']},
    platforms='Win32 (MS Windows)',
    keywords=['xls', 'excel', 'spreadsheet', 'workbook'],
    classifiers=[
        'Development Status :: 4 - Beta',
        'Environment :: Win32 (MS Windows)',
        'Operating System :: Microsoft :: Windows',
        'Programming Language :: Python',
        'Programming Language :: Python :: 2',
        'Programming Language :: Python :: 3',
        'Topic :: Office/Business :: Financial :: Spreadsheet',
        'License :: OSI Approved :: BSD License']
)
