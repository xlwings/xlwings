import os
import json
import binascii
import datetime as dt

from ..utils import read_config_sheet
from .. import LicenseError, Book, __version__, xlplatform

try:
    from cryptography.fernet import Fernet, InvalidToken
except ImportError as e:
    raise ImportError("Couldn't find 'cryptography', a dependency of xlwings PRO.") from None


class LicenseHandler:
    @staticmethod
    def get_cipher():
        try:
            return Fernet(os.getenv('XLWINGS_LICENSE_KEY_SECRET'))
        except (TypeError, ValueError):
            raise LicenseError("Couldn't validate license key.") from None

    @staticmethod
    def get_license():
        # Sheet config (only used by RunPython, UDFs use env var)
        try:
            sheet_license_key = read_config_sheet(Book.caller()).get('LICENSE_KEY')
            if sheet_license_key:
                return sheet_license_key
        except:
            pass
        # User config file
        config_file = xlplatform.USER_CONFIG_FILE
        if os.path.exists(config_file):
            with open(config_file, 'r') as f:
                config = f.readlines()
            key = None
            for line in config:
                if line.split(',')[0] == '"LICENSE_KEY"':
                    key = line.split(',')[1].strip()[1:-1]
            if key:
                return key
        # Env Var - also used if LICENSE_KEY is in config sheet and called via UDF
        if os.getenv('XLWINGS_LICENSE_KEY'):
            return os.environ['XLWINGS_LICENSE_KEY']
        raise LicenseError("Couldn't find a license key.")

    @staticmethod
    def validate_license(product, license_type=None):
        cipher_suite = LicenseHandler.get_cipher()
        key = LicenseHandler.get_license()
        try:
            license_info = json.loads(cipher_suite.decrypt(key.encode()).decode())
        except (binascii.Error, InvalidToken):
            raise LicenseError('Invalid license key.') from None
        try:
            if license_type == 'developer' and license_info['license_type'] != 'developer':
                raise LicenseError('You need a developer license for this action.')
        except KeyError:
            raise LicenseError('You need a developer license for this action.') from None
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
        if 'version' in license_info.keys() and license_info['version'] != __version__:
            raise LicenseError('Your license key is only valid for xlwings v{0}'.format(license_info['version'])) from None

    @staticmethod
    def create_deploy_key():
        LicenseHandler.validate_license('pro', license_type='developer')
        cipher_suite = LicenseHandler.get_cipher()
        license_dict = json.dumps({'version': __version__,
                                   'products': ['pro', 'reports'],
                                   'valid_until': '2999-12-31',
                                   'license_type': 'deploy_key'}).encode()
        return cipher_suite.encrypt(license_dict).decode()
