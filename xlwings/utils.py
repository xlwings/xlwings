import os
import re
import sys
import tempfile
import datetime as dt
import traceback
from functools import total_ordering, lru_cache
from pathlib import Path

try:
    import numpy as np
except ImportError:
    np = None

try:
    import matplotlib as mpl
    import matplotlib.figure
except ImportError:
    mpl = None

try:
    import plotly.graph_objects as plotly_go
except ImportError:
    plotly_go = None

import xlwings

missing = object()


def int_to_rgb(number):
    """Given an integer, return the rgb"""
    number = int(number)
    r = number % 256
    g = (number // 256) % 256
    b = (number // (256 * 256)) % 256
    return r, g, b


def rgb_to_int(rgb):
    """Given an rgb, return an int"""
    return rgb[0] + (rgb[1] * 256) + (rgb[2] * 256 * 256)


def get_duplicates(seq):
    seen = set()
    duplicates = set(x for x in seq if x in seen or seen.add(x))
    return duplicates


def np_datetime_to_datetime(np_datetime):
    ts = (np_datetime - np.datetime64('1970-01-01T00:00:00Z')) / np.timedelta64(1, 's')
    dt_datetime = dt.datetime.utcfromtimestamp(ts)
    return dt_datetime


ALPHABET = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"


def col_name(i):
    i -= 1
    if i < 0:
        raise IndexError(i)
    elif i < 26:
        return ALPHABET[i]
    elif i < 702:
        i -= 26
        return ALPHABET[i//26] + ALPHABET[i%26]
    elif i < 16384:
        i -= 702
        return ALPHABET[i//676] + ALPHABET[i//26%26] + ALPHABET[i%26]
    else:
        raise IndexError(i)


class VBAWriter:

    MAX_VBA_LINE_LENGTH = 1024
    VBA_LINE_SPLIT = ' _\n'
    MAX_VBA_SPLITTED_LINE_LENGTH = MAX_VBA_LINE_LENGTH - len(VBA_LINE_SPLIT)

    class Block:
        def __init__(self, writer, start):
            self.writer = writer
            self.start = start

        def __enter__(self):
            self.writer.writeln(self.start)
            self.writer._indent += 1

        def __exit__(self, exc_type, exc_val, exc_tb):
            self.writer._indent -= 1

    def __init__(self, f):
        self.f = f
        self._indent = 0
        self._freshline = True

    def block(self, template, **kwargs):
        return VBAWriter.Block(self, template.format(**kwargs))

    def start_block(self, template, **kwargs):
        self.writeln(template, **kwargs)
        self._indent += 1

    def end_block(self, template, **kwargs):
        self.writeln(template, **kwargs)
        self._indent -= 1

    def write(self, template, **kwargs):
        if kwargs:
            template = template.format(**kwargs)
        if self._freshline:
            template = ('    ' * self._indent) + template
            self._freshline = False
        self.write_vba_line(template)
        if template[-1] == '\n':
            self._freshline = True

    def write_label(self, label):
        self._indent -= 1
        self.write(label + ':\n')
        self._indent += 1

    def writeln(self, template, **kwargs):
        self.write(template + '\n', **kwargs)

    def write_vba_line(self, vba_line):
        if len(vba_line) > VBAWriter.MAX_VBA_LINE_LENGTH:
            separator_index = VBAWriter.get_separator_index(vba_line)
            self.f.write(vba_line[:separator_index] + VBAWriter.VBA_LINE_SPLIT)
            self.write_vba_line(vba_line[separator_index:])
        else:
            self.f.write(vba_line)

    @classmethod
    def get_separator_index(cls, vba_line):
        for index in range(cls.MAX_VBA_SPLITTED_LINE_LENGTH, 0, -1):
            if ' ' == vba_line[index]:
                return index
        return cls.MAX_VBA_SPLITTED_LINE_LENGTH  # Best effort: split string at the maximum possible length


def try_parse_int(x):
    try:
        return int(x)
    except ValueError:
        return x


@total_ordering
class VersionNumber:

    def __init__(self, s):
        self.value = tuple(map(try_parse_int, s.split(".")))

    @property
    def major(self):
        return self.value[0]

    @property
    def minor(self):
        return self.value[1] if len(self.value) > 1 else None

    def __str__(self):
        return ".".join(map(str, self.value))

    def __repr__(self):
        return "%s(%s)" % (self.__class__.__name__, repr(str(self)))

    def __eq__(self, other):
        if isinstance(other, VersionNumber):
            return self.value == other.value
        elif isinstance(other, str):
            return self.value == VersionNumber(other).value
        elif isinstance(other, tuple):
            return self.value[:len(other)] == other
        elif isinstance(other, int):
            return self.major == other
        else:
            return False

    def __lt__(self, other):
        if isinstance(other, VersionNumber):
            return self.value < other.value
        elif isinstance(other, str):
            return self.value < VersionNumber(other).value
        elif isinstance(other, tuple):
            return self.value[:len(other)] < other
        elif isinstance(other, int):
            return self.major < other
        else:
            raise TypeError("Cannot compare other object with version number")


def process_image(image, width, height):
    image = fspath(image)
    if isinstance(image, str):
        return image, width, height
    elif mpl and isinstance(image, mpl.figure.Figure):
        image_type = 'mpl'
    elif plotly_go and isinstance(image, plotly_go.Figure) and xlwings.PRO:
        image_type = 'plotly'
    else:
        raise TypeError("Don't know what to do with that image object")

    temp_dir = os.path.realpath(tempfile.gettempdir())
    filename = os.path.join(temp_dir, 'xlwings_plot.png')

    if image_type == 'mpl':
        canvas = mpl.backends.backend_agg.FigureCanvas(image)
        canvas.draw()
        image.savefig(filename, format='png', bbox_inches='tight', dpi=300)

        if width is None:
            width = image.bbox.bounds[2:][0]

        if height is None:
            height = image.bbox.bounds[2:][1]
    elif image_type == 'plotly':
        image.write_image(filename, width=None, height=None)
    return filename, width, height


def fspath(path):
    """Convert path-like object to string.

    On python <= 3.5 the input argument is always returned unchanged (no support for path-like
    objects available).

    """
    if hasattr(os, 'PathLike') and isinstance(path, os.PathLike):
        return os.fspath(path)
    else:
        return path


def read_config_sheet(book):
    try:
        return book.sheets['xlwings.conf']['A1:B1'].options(dict, expand='down').value
    except:
        # A missing sheet currently produces different errors on mac and win
        return {}


def read_user_config():
    """Returns keys in lowercase of xlwings.conf in the user's home directory"""
    config = {}
    if Path(xlwings.USER_CONFIG_FILE).is_file():
        with open(xlwings.USER_CONFIG_FILE, 'r') as f:
            for line in f:
                values = re.findall(r'"[^"]*"', line)
                if values:
                    config[values[0].strip('"').lower()] = values[1].strip('"')
    return config


@lru_cache(None)
def get_cached_user_config(key):
    return read_user_config().get(key.lower())


def exception(logger, msg, *args):
    if logger.hasHandlers():
        logger.exception(msg, *args)
    else:
        print(msg % args)
        traceback.print_exc()


def chunk(sequence, chunksize):
    for i in range(0, len(sequence), chunksize):
        yield sequence[i:i+chunksize]


def query_yes_no(question, default="yes"):
    """Ask a yes/no question via input() and return their answer.

    "question" is a string that is presented to the user.
    "default" is the presumed answer if the user just hits <Enter>.
            It must be "yes" (the default), "no" or None (meaning
            an answer is required of the user).

    The "answer" return value is True for "yes" or False for "no".

    Licensed under the MIT License
    Copyright by Trent Mick
    https://code.activestate.com/recipes/577058/
    """
    valid = {"yes": True, "y": True, "ye": True, "no": False, "n": False}
    if default is None:
        prompt = " [y/n] "
    elif default == "yes":
        prompt = " [Y/n] "
    elif default == "no":
        prompt = " [y/N] "
    else:
        raise ValueError("invalid default answer: '%s'" % default)

    while True:
        sys.stdout.write(question + prompt)
        choice = input().lower()
        if default is not None and choice == "":
            return valid[default]
        elif choice in valid:
            return valid[choice]
        else:
            sys.stdout.write("Please respond with 'yes' or 'no' " "(or 'y' or 'n').\n")
