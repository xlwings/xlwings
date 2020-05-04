import binascii
import sys
import os
import datetime as dt
import json
from .. import LicenseError

try:
    from cryptography.fernet import Fernet, InvalidToken
except ImportError as e:
    raise ImportError("Couldn't find 'cryptography', a dependency of xlwings PRO.") from None


class LicenseHandler:
    @staticmethod
    def get_license():
        if os.getenv('XLWINGS_LICENSE_KEY'):
            return os.environ['XLWINGS_LICENSE_KEY']
        if sys.platform.startswith('darwin'):
            config_file = os.path.join(os.path.expanduser("~"), 'Library', 'Containers',
                                       'com.microsoft.Excel', 'Data', 'xlwings.conf')
        else:
            config_file = os.path.join(os.path.expanduser("~"), '.xlwings', 'xlwings.conf')
        if not os.path.exists(config_file):
            raise LicenseError("Couldn't find a license key.")
        with open(config_file, 'r') as f:
            config = f.readlines()
        key = None
        for line in config:
            if line.split(',')[0] == '"LICENSE_KEY"':
                key = line.split(',')[1].strip()[1:-1]
        if key:
            return key
        else:
            raise LicenseError("Couldn't find a valid license key.") from None

    @staticmethod
    def validate_license(product):
        try:
            cipher_suite = Fernet(os.getenv('XLWINGS_LICENSE_KEY_SECRET'))
        except ValueError:
            raise LicenseError("Couldn't validate license key.") from None
        key = LicenseHandler.get_license()
        try:
            license_info = json.loads(cipher_suite.decrypt(key.encode()).decode())
        except (binascii.Error, InvalidToken):
            raise LicenseError('Invalid license key.') from None
        if 'valid_until' not in license_info.keys() or 'products' not in license_info.keys():
            raise LicenseError('Invalid license key format.') from None
        if 'valid_until' not in license_info.keys() or 'products' not in license_info.keys():
            raise LicenseError('Invalid license key format.') from None
        license_valid_until = dt.datetime.strptime(license_info['valid_until'], '%Y-%m-%d').date()
        if dt.date.today() > license_valid_until:
            raise LicenseError('Your license expired on {}.'.format(license_valid_until.strftime("%Y-%m-%d"))) from None
        if product not in license_info['products']:
            if product == 'pro':
                raise LicenseError('Invalid license key for xlwings PRO.') from None
            elif product == 'reports':
                raise LicenseError('Your license is not valid for the xlwings reports add-on.') from None
