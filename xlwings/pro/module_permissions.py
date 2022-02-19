"""
Required Notice: Copyright (C) Zoomer Analytics GmbH.

xlwings PRO is dual-licensed under one of the following licenses:

* PolyForm Noncommercial License 1.0.0 (for noncommercial use):
  https://polyformproject.org/licenses/noncommercial/1.0.0
* xlwings PRO License (for commercial use):
  https://github.com/xlwings/xlwings/blob/main/LICENSE_PRO.txt

Commercial licenses can be purchased at https://www.xlwings.org
"""

import re
import sys
import hashlib
import socket
from functools import lru_cache
from pathlib import Path

try:
    import requests
except ImportError:
    requests = None

from . import LicenseHandler
from .. import XlwingsError
from ..utils import read_user_config


@lru_cache(None)
def verify_execute_permission(command=None, module_names=None):
    LicenseHandler.validate_license("permissioning")
    if command:
        assert not module_names
        if re.compile(r"from .* import .*").search(command):
            raise XlwingsError("Can't verify 'from x import y' imports.")
        module_names = re.findall(r"import ([^;]*)", command)
    elif module_names:
        assert not command
    else:
        raise ValueError("You must either provide command or module_names!")
    file_names = [module + ".py" for module in module_names]
    file_hashes = {}
    for fn in file_names:
        for path in sys.path:
            # Can't use pkgutil or importlib as they may import, i.e. run the module
            if (Path(path) / fn).is_file():
                with open(Path(path) / fn, "rb") as f:
                    content = f.read()
                file_hashes[fn] = hashlib.sha256(content).hexdigest()
                break
        if fn not in file_hashes:
            raise FileNotFoundError(f"Couldn't find {fn}")

    config = read_user_config()
    method = config.get("permission_check_method", "GET").upper()

    if method == "GET":
        response = requests.get(config["permission_check_url"], timeout=10)
        if response.status_code != 200:
            raise XlwingsError(
                f"Failed to connect to permission server. Error {response.status_code}."
            )
        response = response.json()
        checked_files = []
        for file_name in file_names:
            for module in response["modules"]:
                if file_name == module["file_name"]:
                    correct_sha256 = file_hashes[file_name] == module["sha256"]
                    permitted_machine = (
                        module["machine_names"] == "*"
                        or "*" in module["machine_names"]
                        or socket.gethostname() in module["machine_names"]
                    )
                    if correct_sha256 and permitted_machine:
                        checked_files.append(file_name)
                        break
                    else:
                        raise XlwingsError(
                            f"Failed to get permission for the following file: "
                            f"{file_name}"
                        )
        missing_permissions = set(file_names).difference(set(checked_files))
        if missing_permissions:
            raise XlwingsError(
                f"Failed to get permission for the following file(s): "
                f"{', '.join(missing_permissions)}"
            )
    elif method == "POST":
        payload = {"machine_name": socket.gethostname(), "modules": []}
        for file_name, sha256 in file_hashes.items():
            payload["modules"].append({"file_name": file_name, "sha256": sha256})
        headers = (
            {"Authorization": config["permission_check_authorization"]}
            if config.get("permission_check_authorization")
            else None
        )
        response = requests.post(
            config["permission_check_url"],
            json=payload,
            timeout=10,
            headers=headers,
        )

        if response.status_code == 200:
            return True
        else:
            raise XlwingsError(
                f"Failed to get permission for the following file(s): "
                f"{', '.join(file_names)}. Error {response.status_code}."
            )
    else:
        raise ValueError("PERMISSION_CHECK_URL must be either GET or POST.")
