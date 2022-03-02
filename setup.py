import os
import sys
import re
import glob
from setuptools import setup, find_packages

# long_description: Take from README file
with open(os.path.join(os.path.dirname(__file__), "README.rst")) as f:
    readme = f.read()

# Version Number
with open(os.path.join(os.path.dirname(__file__), "xlwings", "__init__.py")) as f:
    version = re.compile(r'.*__version__ = "(.*?)"', re.S).match(f.read()).group(1)

# Dependencies
data_files = []
install_requires = []
if os.environ.get("READTHEDOCS", None) == "True" or os.environ.get(
    "XLWINGS_NO_DEPS"
) in ["1", "True", "true"]:
    # Don't add any further dependencies. Instead of using an env var,
    # you could also run: pip install xlwings --no-deps
    # but when running "pip install -r requirements.txt --no-deps" this would be
    # applied to all packages, which may not be what you want in case the
    # sub-dependencies are not pinned
    pass
elif sys.platform.startswith("win"):
    if sys.version_info[:2] >= (3, 7):
        pywin32 = "pywin32 >= 224"
    else:
        pywin32 = "pywin32"
    install_requires += [pywin32]
    # This places dlls next to python.exe for standard setup
    # and in the parent folder for virtualenv
    data_files += [("", glob.glob("xlwings??-*.dll"))]
elif sys.platform.startswith("darwin"):
    install_requires += ["psutil >= 2.0.0", "appscript >= 1.0.1"]
    # This has been broken ever since pip builds a wheel for installing as this uses
    # an absolute path. Relative paths as in the case of Windows are still supported.
    # These days taken care of by "xlwings addin install" or "xlwings runpython install"
    data_files = [
        (
            os.path.expanduser("~")
            + "/Library/Application Scripts/com.microsoft.Excel",
            [f"xlwings/xlwings-{version}.applescript"],
        )
    ]
else:
    pass

extras_require = {
    "pro": ["cryptography"],
    "reports": ["cryptography", "Jinja2", "pdfrw"],
    "all": [
        "cryptography",
        "Jinja2",
        "pandas",
        "matplotlib",
        "plotly",
        "flask",
        "requests",
        "pdfrw",
    ],
}

setup(
    name="xlwings",
    version=version,
    url="https://www.xlwings.org",
    project_urls={
        "Source": "https://github.com/xlwings/xlwings",
        "Documentation": "https://docs.xlwings.org",
    },
    license="BSD 3-clause",
    author="Zoomer Analytics LLC",
    author_email="felix.zumstein@zoomeranalytics.com",
    description="Make Excel fly: Interact with Excel from Python and vice versa.",
    long_description=readme,
    data_files=data_files,
    packages=find_packages(
        exclude=(
            "tests",
            "tests.*",
        )
    ),
    package_data={
        "xlwings": [
            "xlwings.bas",
            "Dictionary.cls",
            "*.xlsm",
            "*.xlam",
            "*.applescript",
            "addin/xlwings.xlam",
            "addin/xlwings_unprotected.xlam",
            "js/xlwings.*",
            "quickstart_fastapi/*.*",
        ]
    },
    keywords=["xls", "excel", "spreadsheet", "workbook", "vba", "macro"],
    install_requires=install_requires,
    extras_require=extras_require,
    entry_points={
        "console_scripts": ["xlwings=xlwings.cli:main"],
    },
    classifiers=[
        "Development Status :: 4 - Beta",
        "Operating System :: OS Independent",
        "Programming Language :: Python",
        "Programming Language :: Python :: 3",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Topic :: Office/Business :: Financial :: Spreadsheet",
        "License :: OSI Approved :: BSD License",
    ],
    python_requires=">=3.6",
)
