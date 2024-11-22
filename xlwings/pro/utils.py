"""
Required Notice: Copyright (C) Zoomer Analytics GmbH.

xlwings PRO is dual-licensed under one of the following licenses:

* PolyForm Noncommercial License 1.0.0 (for noncommercial use):
  https://polyformproject.org/licenses/noncommercial/1.0.0
* xlwings PRO License (for commercial use):
  https://github.com/xlwings/xlwings/blob/main/LICENSE_PRO.txt

Commercial licenses can be purchased at https://www.xlwings.org
"""

import atexit
import base64
import binascii
import datetime as dt
import glob
import hashlib
import hmac
import json
import os
import shutil
import tempfile
import time
import warnings
from functools import lru_cache

from packaging.version import Version

import xlwings  # To prevent circular imports

from ..utils import read_config_sheet

try:
    from cryptography.fernet import Fernet, InvalidToken
except ImportError:
    Fernet = None
    InvalidToken = None


class LicenseHandler:
    @staticmethod
    def get_cipher():
        try:
            return Fernet(os.getenv("XLWINGS_LICENSE_KEY_SECRET"))
        except (TypeError, ValueError):
            raise xlwings.LicenseError(
                "Couldn't validate xlwings license key."
            ) from None

    @staticmethod
    def get_license():
        # Env Var - also used if LICENSE_KEY is in config sheet and called via UDF
        if os.getenv("XLWINGS_LICENSE_KEY"):
            return os.environ["XLWINGS_LICENSE_KEY"]
        # Sheet config (only used by RunPython, UDFs use env var)
        try:
            sheet_license_key = read_config_sheet(xlwings.Book.caller()).get(
                "LICENSE_KEY"
            )
            if sheet_license_key:
                return sheet_license_key
        except:  # noqa: E722
            pass
        # User config file
        config_file = xlwings.USER_CONFIG_FILE
        if os.path.exists(config_file):
            with open(config_file, "r") as f:
                config = f.readlines()
            key = None
            for line in config:
                if line.split(",")[0] == '"LICENSE_KEY"':
                    key = line.split(",")[1].strip()[1:-1]
            if key:
                return key
        raise xlwings.LicenseError("Couldn't find an xlwings license key.")

    @staticmethod
    @lru_cache()
    def validate_license(product, license_type=None):
        key = LicenseHandler.get_license()
        if key == "noncommercial":
            return {"license_type": "noncommercial"}
        if key.startswith("gA") and not Fernet:
            # Legacy up to 0.27.12
            raise ImportError(
                "You are using a legacy xlwings license key that requires the "
                "'cryptography' package. Either install it via 'pip install "
                "cryptography' or contact us for a new license key that doesn't depend "
                "on cryptography."
            ) from None
        elif key.startswith("gA"):
            cipher_suite = LicenseHandler.get_cipher()
            try:
                license_info = json.loads(cipher_suite.decrypt(key.encode()).decode())
            except (binascii.Error, InvalidToken):
                raise xlwings.LicenseError("Invalid xlwings license key.") from None
        else:
            signature = hmac.new(
                os.getenv("XLWINGS_LICENSE_KEY_SECRET").encode(),
                key[:-5].encode(),
                hashlib.sha256,
            ).hexdigest()
            if signature[:5] != key[-5:]:
                raise xlwings.LicenseError("Invalid xlwings license key.") from None
            else:
                try:
                    license_info = json.loads(
                        base64.urlsafe_b64decode(key[:-5]).decode()
                    )
                except:  # noqa: E722
                    raise xlwings.LicenseError("Invalid xlwings license key.") from None
        try:
            if (
                license_type == "developer"
                and license_info["license_type"] != "developer"
            ):
                raise xlwings.LicenseError(
                    "You need a paid xlwings license key for this action."
                )
        except KeyError:
            raise xlwings.LicenseError(
                "You need a paid xlwings license key for this action."
            ) from None
        if (
            "valid_until" not in license_info.keys()
            or "products" not in license_info.keys()
        ):
            raise xlwings.LicenseError("Invalid xlwings license key format.") from None
        license_valid_until = dt.datetime.strptime(
            license_info["valid_until"], "%Y-%m-%d"
        ).date()
        if dt.date.today() > license_valid_until:
            raise xlwings.LicenseError(
                "Your xlwings license expired on {}.".format(
                    license_valid_until.strftime("%Y-%m-%d")
                )
            ) from None
        if product not in license_info["products"]:
            raise xlwings.LicenseError(
                f"Your xlwings license key isn't valid for the '{product}' "
                "functionality."
            ) from None
        if (
            "version" in license_info.keys()
            and license_info["version"] != xlwings.__version__
        ):
            raise xlwings.LicenseError(
                "Your xlwings deploy key is only valid for v{0}. To use a different "
                "version of xlwings, re-release your tool or generate a new deploy key "
                "via 'xlwings license deploy'.".format(license_info["version"])
            ) from None
        if (license_valid_until - dt.date.today()) < dt.timedelta(days=30):
            warnings.warn(
                f"Your xlwings license key expires in "
                f"{(license_valid_until - dt.date.today()).days} days."
            )
        return license_info

    @staticmethod
    def create_deploy_key(version=None):
        license_info = LicenseHandler.validate_license("pro", license_type="developer")
        if license_info["license_type"] == "noncommercial":
            return "noncommercial"

        if version and Version(version) <= Version(xlwings.__version__):
            license_version = version
        else:
            license_version = xlwings.__version__

        license_dict = json.dumps(
            {
                "version": license_version,
                "products": license_info["products"],
                "valid_until": "2999-12-31",
                "license_type": "deploy_key",
            }
        ).encode()

        if LicenseHandler.get_license().startswith("gA"):
            # Legacy
            cipher_suite = LicenseHandler.get_cipher()
            return cipher_suite.encrypt(license_dict).decode()
        else:
            body = base64.urlsafe_b64encode(license_dict)
            signature = hmac.new(
                os.getenv("XLWINGS_LICENSE_KEY_SECRET").encode(), body, hashlib.sha256
            ).hexdigest()
            return f"{body.decode()}{signature[:5]}"


@lru_cache()
def get_embedded_code_temp_dir():
    tmp_base_path = os.path.join(tempfile.gettempdir(), "xlwings")
    os.makedirs(tmp_base_path, exist_ok=True)
    try:
        # HACK: Clean up directories that are older than 30 days
        # This should be done in the C++ part when the Python process is killed
        for subdir in glob.glob(tmp_base_path + "/*/"):
            if os.path.getmtime(subdir) < time.time() - 30 * 86400:
                shutil.rmtree(subdir, ignore_errors=True)
    except Exception:
        pass  # we don't care if it fails
    tempdir = tempfile.mkdtemp(dir=tmp_base_path)
    # This only works for RunPython calls running outside the COM server
    atexit.register(shutil.rmtree, tempdir)
    return tempdir
