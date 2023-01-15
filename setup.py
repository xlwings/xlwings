"""
Windows builds currently rely on setup.py instead of pyproject.toml/maturin as long
as the dlls are distributed as data_files
"""
import glob
import os
import re
import sys

from setuptools import find_packages, setup

# long_description: Take from README file
with open(os.path.join(os.path.dirname(__file__), "README.rst")) as f:
    readme = f.read()

# Version Number
with open(os.path.join(os.path.dirname(__file__), "xlwings", "__init__.py")) as f:
    version = re.compile(r'.*__version__ = "(.*?)"', re.S).match(f.read()).group(1)

# Dependencies
data_files = []
install_requires = [
    "pywin32 >= 224;platform_system=='Windows'",
    "psutil >= 2.0.0;platform_system=='Darwin'",
    "appscript >= 1.0.1;platform_system=='Darwin'",
]
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
    # This places dlls next to python.exe for standard setup
    # and in the parent folder for virtualenv
    data_files += [("", glob.glob("xlwings??-*.dll"))]
else:
    pass

extras_require = {
    "reports": ["Jinja2", "pdfrw"],
    "all": [
        "Jinja2",
        "pandas",
        "matplotlib",
        "plotly",
        "flask",
        "requests",
        "pdfrw",
    ],
}

if os.getenv("BUILD_RUST", "0") == "1":
    from setuptools_rust import Binding, RustExtension

    rust_extensions = [
        RustExtension(
            "xlwings.xlwingslib",
            binding=Binding.PyO3,
            path="./Cargo.toml",
        )
    ]
else:
    rust_extensions = []

setup(
    name="xlwings",
    version=version,
    rust_extensions=rust_extensions,
    zip_safe=False,  # Rust extensions are not zip safe
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
            "xlwings_custom_addin.bas",
            "*.xlsm",
            "*.xlam",
            "*.applescript",
            "addin/xlwings.xlam",
            "addin/*.cls",
            "addin/WebHelpers.bas",
            "js/xlwings.*",
            "quickstart_fastapi/*.*",
            "html/*.*",
        ],
        "src": [
            "src",
            "Cargo.*",
        ],
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
    python_requires=">=3.7",
)
