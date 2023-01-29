"""
This is run by GH Actions to make sure the version corresponds to the release version
"""

import os
from pathlib import Path

this_dir = Path(__file__).resolve().parent
if os.environ["GITHUB_REF"].startswith("refs/tags"):
    version = os.environ["GITHUB_REF"][10:]
else:
    raise Exception("Didn't find Git tag!")

xlwingsjs = this_dir / "dist" / "xlwings.js"
xlwingsminjs = this_dir / "dist" / "xlwings.min.js"

if f'version = "{version}"' not in xlwingsjs.read_text():
    raise Exception("Didn't find expected version in xlwings.js!")

if f'b.version="{version}"' not in xlwingsminjs.read_text():
    raise Exception("Didn't find expected version in xlwings.min.js!")
