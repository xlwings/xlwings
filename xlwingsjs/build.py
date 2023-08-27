import argparse
import os
import shutil
import subprocess
from pathlib import Path
from shlex import split
from textwrap import dedent

this_dir = Path(__file__).resolve().parent
os.chdir(this_dir)

target_dir = "dist"
webpack_config = "webpack.config.js"
source = "src/xlwings.ts"


def webpack():
    return subprocess.run(
        split("npx webpack"),
        stdout=subprocess.PIPE,
        stderr=subprocess.STDOUT,
        encoding="utf-8",
    )


def prepend_license(path):
    content = Path(path).read_text()
    content = (
        dedent(
            """\
            /**
            * Copyright (C) 2014 - present, Zoomer Analytics GmbH. All rights reserved.
            * Licensed under BSD-3-Clause license, see: https://docs.xlwings.org/en/stable/license.html
            *
            * This file also contains code from core-js
            * Copyright (C) 2014-2023 Denis Pushkarev, Licensed under MIT license, see https://raw.githubusercontent.com/zloirock/core-js/master/LICENSE
            * This file also contains code from Webpack
            * Copyright (C) JS Foundation and other contributors, Licensed under MIT license, see https://raw.githubusercontent.com/webpack/webpack/main/LICENSE
            */
            """
        )
        + content
    )
    Path(path).write_text(content)


def build(version):
    # Clear
    if Path(target_dir).exists():
        shutil.rmtree(target_dir)

    # Version
    content = Path(source).read_text()
    content = content.replace('version = "dev"', f'version = "{version}"')
    Path(source).write_text(content)

    # Minified build
    print(webpack())
    prepend_license(f"{target_dir}/xlwings.min.js")

    # Non-minified build
    content = Path(webpack_config).read_text()
    content = content.replace('mode: "production"', 'mode: "development"').replace(
        'filename: "xlwings.min.js"', 'filename: "xlwings.js"'
    )
    Path(webpack_config).write_text(content)
    print(webpack())
    prepend_license(f"{target_dir}/xlwings.js")

    # Reset version
    content = Path(source).read_text()
    content = content.replace(f'version = "{version}"', 'version = "dev"')
    Path(source).write_text(content)

    # Reset webpack.config.js
    content = Path(webpack_config).read_text()
    content = content.replace('mode: "development"', 'mode: "production"').replace(
        'filename: "xlwings.js"', 'filename: "xlwings.min.js"'
    )
    Path(webpack_config).write_text(content)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--version", required=True)
    args = parser.parse_args()
    build(args.version)
