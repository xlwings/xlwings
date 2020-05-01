import binascii
import sys
import os
import datetime as dt
import json

try:
    from cryptography.fernet import Fernet, InvalidToken
except ImportError as e:
    sys.exit("Couldn't find 'cryptography', a dependency of xlwings PRO. "
             "Details: {0}.".format(repr(e)))


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
            sys.exit("Couldn't find a license key.")
        with open(config_file, 'r') as f:
            config = f.readlines()
        key = None
        for line in config:
            if line.split(',')[0] == '"LICENSE_KEY"':
                key = line.split(',')[1].strip()[1:-1]
        if key:
            return key
        else:
            sys.exit("Couldn't find a valid license key.")

    @staticmethod
    def validate_license(product):
        cipher_suite = Fernet(os.getenv('LICENSE_KEY_SECRET'))
        key = LicenseHandler.get_license()
        try:
            license_info = json.loads(cipher_suite.decrypt(key.encode()).decode())
        except (binascii.Error, InvalidToken):
            sys.exit('Invalid license key.')
        if 'valid_until' not in license_info.keys() or 'products' not in license_info.keys():
            sys.exit('Invalid license key format.')
        if 'valid_until' not in license_info.keys() or 'products' not in license_info.keys():
            sys.exit('Invalid license key format.')
        license_valid_until = dt.datetime.strptime(license_info['valid_until'], '%Y-%m-%d').date()
        if dt.date.today() > license_valid_until:
            sys.exit('Your license expired on {}.'.format(license_valid_until.strftime("%Y-%m-%d")))
        if product not in license_info['products']:
            if product == 'pro':
                sys.exit('Invalid license key for xlwings PRO.')
            elif product == 'reports':
                sys.exit('Your license is not valid for the xlwings reports add-on.')
