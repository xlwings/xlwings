import os
import re
import sys
import uuid
import tempfile
import subprocess
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
    import matplotlib.pyplot as plt
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


def hex_to_rgb(color):
    color = color[1:] if color.startswith("#") else color
    return tuple(int(color[i : i + 2], 16) for i in (0, 2, 4))


def rgb_to_hex(r, g, b):
    return f"#{r:02x}{g:02x}{b:02x}"


def get_duplicates(seq):
    seen = set()
    duplicates = set(x for x in seq if x in seen or seen.add(x))
    return duplicates


def np_datetime_to_datetime(np_datetime):
    ts = (np_datetime - np.datetime64("1970-01-01T00:00:00Z")) / np.timedelta64(1, "s")
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
        return ALPHABET[i // 26] + ALPHABET[i % 26]
    elif i < 16384:
        i -= 702
        return ALPHABET[i // 676] + ALPHABET[i // 26 % 26] + ALPHABET[i % 26]
    else:
        raise IndexError(i)


def address_to_index_tuple(address):
    """
    Based on a function from XlsxWriter, which is distributed under the following
    BSD 2-Clause License:

    Copyright (c) 2013-2021, John McNamara <jmcnamara@cpan.org>
    All rights reserved.

    Redistribution and use in source and binary forms, with or without
    modification, are permitted provided that the following conditions are met:

    1. Redistributions of source code must retain the above copyright notice, this
       list of conditions and the following disclaimer.

    2. Redistributions in binary form must reproduce the above copyright notice,
       this list of conditions and the following disclaimer in the documentation
       and/or other materials provided with the distribution.

    THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
    AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
    IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
    DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE
    FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL
    DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR
    SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER
    CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY,
    OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE
    OF THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
    """
    re_range_parts = re.compile(r"(\$?)([A-Z]{1,3})(\$?)(\d+)")
    match = re_range_parts.match(address)
    col_str = match.group(2)
    row_str = match.group(4)

    # Convert base26 column string to number
    expn = 0
    col = 0
    for char in reversed(col_str):
        col += (ord(char) - ord("A") + 1) * (26**expn)
        expn += 1

    return int(row_str), col


class VBAWriter:

    MAX_VBA_LINE_LENGTH = 1024
    VBA_LINE_SPLIT = " _\n"
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
            template = ("    " * self._indent) + template
            self._freshline = False
        self.write_vba_line(template)
        if template[-1] == "\n":
            self._freshline = True

    def write_label(self, label):
        self._indent -= 1
        self.write(label + ":\n")
        self._indent += 1

    def writeln(self, template, **kwargs):
        self.write(template + "\n", **kwargs)

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
            if " " == vba_line[index]:
                return index
        return (
            cls.MAX_VBA_SPLITTED_LINE_LENGTH
        )  # Best effort: split string at the maximum possible length


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
            return self.value[: len(other)] == other
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
            return self.value[: len(other)] < other
        elif isinstance(other, int):
            return self.major < other
        else:
            raise TypeError("Cannot compare other object with version number")


def process_image(image, format):
    """Returns filename and is_temp_file"""
    image = fspath(image)
    if isinstance(image, str):
        return image, False
    elif mpl and isinstance(image, mpl.figure.Figure):
        image_type = "mpl"
    elif plotly_go and isinstance(image, plotly_go.Figure):
        image_type = "plotly"
    else:
        raise TypeError("Don't know what to do with that image object")

    if format == "vector":
        if sys.platform.startswith("darwin"):
            format = "pdf"
        else:
            format = "svg"

    temp_dir = os.path.realpath(tempfile.gettempdir())
    filename = os.path.join(temp_dir, str(uuid.uuid4()) + "." + format)

    if image_type == "mpl":
        canvas = mpl.backends.backend_agg.FigureCanvas(image)
        canvas.draw()
        image.savefig(filename, bbox_inches="tight", dpi=300)
        plt.close(image)
    elif image_type == "plotly":
        image.write_image(filename)
    return filename, True


def fspath(path):
    """Convert path-like object to string.

    On python <= 3.5 the input argument is always returned unchanged (no support for
    path-like objects available). TODO: can be removed as 3.5 no longer supported.

    """
    if hasattr(os, "PathLike") and isinstance(path, os.PathLike):
        return os.fspath(path)
    else:
        return path


def read_config_sheet(book):
    try:
        return book.sheets["xlwings.conf"]["A1:B1"].options(dict, expand="down").value
    except:
        # A missing sheet currently produces different errors on mac and win
        return {}


def read_user_config():
    """Returns keys in lowercase of xlwings.conf in the user's home directory"""
    config = {}
    if Path(xlwings.USER_CONFIG_FILE).is_file():
        with open(xlwings.USER_CONFIG_FILE, "r") as f:
            for line in f:
                values = re.findall(r'"[^"]*"', line)
                if values:
                    config[values[0].strip('"').lower()] = os.path.expandvars(
                        values[1].strip('"')
                    )
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
        yield sequence[i : i + chunksize]


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


def prepare_sys_path(args_string):
    """Called from Excel to prepend the default paths and those from the PYTHONPATH
    setting to sys.path. While RunPython could use Book.caller(), the UDF server can't,
    as this runs before VBA can push the ActiveWorkbook over. UDFs also can't interact
    with the book object in general as Excel is busy during the function call and so
    won't allow you to read out the config sheet, for example. Before 0.24.9,
    these manipulations were handled in VBA, but couldn't handle SharePoint.
    """
    args = os.path.normcase(os.path.expandvars(args_string)).split(";")
    # Not sure, if we really need normcase,
    # but on Windows it replaces "/" with "\", so let's revert that
    active_fullname = args[0].replace("\\", "/")
    this_fullname = args[1].replace("\\", "/")
    paths = []
    for fullname in [active_fullname, this_fullname]:
        if not fullname:
            continue
        elif "://" in fullname:
            fullname = Path(
                fullname_url_to_local_path(
                    url=fullname,
                    sheet_onedrive_consumer_config=args[2],
                    sheet_onedrive_commercial_config=args[3],
                    sheet_sharepoint_config=args[4],
                )
            )
        else:
            fullname = Path(fullname)
        paths += [str(fullname.parent), str(fullname.with_suffix(".zip"))]

    if args[5:]:
        paths += args[5:]

    sys.path[0:0] = list(set(paths))


@lru_cache(None)
def fullname_url_to_local_path(
    url,
    sheet_onedrive_consumer_config=None,
    sheet_onedrive_commercial_config=None,
    sheet_sharepoint_config=None,
):
    """
    When AutoSave is enabled in Excel with either OneDrive or SharePoint, VBA/COM's
    Workbook.FullName turns into a URL without any possibilities to get the local file
    path. While OneDrive and OneDrive for Business make it easy enough to derive the
    local path from the URL, SharePoint allows to define the "Site name" and "Site
    address" independently from each other with the former ending up in the local folder
    path and the latter in the FullName URL. Adding to the complexity: (1) When the site
    name contains spaces, they will be stripped out from the URL and (2) you can sync a
    subfolder directly (this, at least, works when you have a single folder at the
    SharePoint's Document root), which results in skipping a folder level locally when
    compared to the online/URL version. And (3) the OneDriveCommercial env var sometimes
    seems to actually point to the local SharePoint folder.

    Parameters
    ----------
    url : str
        URL as returned by VBA's FullName

    sheet_onedrive_consumer_config : str
        Optional Path to the local OneDrive (Personal) as defined in the Workbook's
        config sheet

    sheet_onedrive_commercial_config : str
        Optional Path to the local OneDrive for Business as defined in the Workbook's
        config sheet

    sheet_sharepoint_config : str
        Optional Path to the local SharePoint drive as defined in the Workbook's config
        sheet
    """
    # Directory config files can't be used
    # since the whole purpose of this exercise is to find out about a book's dir
    onedrive_consumer_config_name = (
        "ONEDRIVE_CONSUMER_WIN"
        if sys.platform.startswith("win")
        else "ONEDRIVE_CONSUMER_MAC"
    )
    onedrive_commercial_config_name = (
        "ONEDRIVE_COMMERCIAL_WIN"
        if sys.platform.startswith("win")
        else "ONEDRIVE_COMMERCIAL_MAC"
    )
    sharepoint_config_name = (
        "SHAREPOINT_WIN" if sys.platform.startswith("win") else "SHAREPOINT_MAC"
    )
    if sheet_onedrive_consumer_config is not None:
        sheet_onedrive_consumer_config = os.path.expandvars(
            sheet_onedrive_consumer_config
        )
    if sheet_onedrive_commercial_config is not None:
        sheet_onedrive_commercial_config = os.path.expandvars(
            sheet_onedrive_commercial_config
        )
    if sheet_sharepoint_config is not None:
        sheet_sharepoint_config = os.path.expandvars(sheet_sharepoint_config)
    onedrive_consumer_config = sheet_onedrive_consumer_config or read_user_config().get(
        onedrive_consumer_config_name.lower()
    )
    onedrive_commercial_config = (
        sheet_onedrive_commercial_config
        or read_user_config().get(onedrive_commercial_config_name.lower())
    )
    sharepoint_config = sheet_sharepoint_config or read_user_config().get(
        sharepoint_config_name.lower()
    )

    # OneDrive
    pattern = re.compile(r"https://d.docs.live.net/[^/]*/(.*)")
    match = pattern.match(url)
    if match:
        root = (
            onedrive_consumer_config
            or os.getenv("OneDriveConsumer")
            or os.getenv("OneDrive")
            or str(Path.home() / "OneDrive")
        )
        if not root:
            raise xlwings.XlwingsError(
                f"Couldn't find the local OneDrive folder. Please configure the "
                f"{onedrive_consumer_config_name} setting, see: xlwings.org/error."
            )
        local_path = Path(root) / match.group(1)
        if local_path.is_file():
            return str(local_path)
        else:
            raise xlwings.XlwingsError(
                "Couldn't find your local OneDrive file, see: xlwings.org/error"
            )

    # OneDrive for Business
    pattern = re.compile(r"https://[^-]*-my.sharepoint.[^/]*/[^/]*/[^/]*/[^/]*/(.*)")
    match = pattern.match(url)
    if match:
        root = (
            onedrive_commercial_config
            or os.getenv("OneDriveCommercial")
            or os.getenv("OneDrive")
        )
        if not root:
            raise xlwings.XlwingsError(
                f"Couldn't find the local OneDrive for Business folder. "
                f"Please configure the {onedrive_commercial_config_name} setting, "
                f"see: xlwings.org/error."
            )
        local_path = Path(root) / match.group(1)
        if local_path.is_file():
            return str(local_path)
        else:
            raise xlwings.XlwingsError(
                "Couldn't find your local OneDrive for Business file, "
                "see: xlwings.org/error"
            )

    # SharePoint Online & On-Premises (default top level mapping)
    pattern = re.compile(r"https?://[^/]*/sites/([^/]*)/([^/]*)/(.*)")
    match = pattern.match(url)
    # We're trying to derive the SharePoint root path
    # from the OneDriveCommercial path, if it exists
    root = sharepoint_config or (
        os.getenv("OneDriveCommercial").replace("OneDrive - ", "")
        if os.getenv("OneDriveCommercial")
        else None
    )
    if not root:
        raise xlwings.XlwingsError(
            f"Couldn't find the local SharePoint folder. Please configure the "
            f"{sharepoint_config_name} setting, see: xlwings.org/error."
        )
    if match:
        local_path = Path(root) / f"{match.group(1)} - Documents" / match.group(3)
        if local_path.is_file():
            return str(local_path)
    # SharePoint Online & On-Premises (non-default mapping)
    book_name = url.split("/")[-1]
    local_book_paths = []
    for path in Path(root).rglob("[!~$]*.xls*"):
        if path.name.lower() == book_name.lower():
            local_book_paths.append(path)
    if len(local_book_paths) == 1:
        return str(local_book_paths[0])
    elif len(local_book_paths) == 0:
        raise xlwings.XlwingsError(
            f"Couldn't find your SharePoint file locally, see: xlwings.org/error"
        )
    else:
        raise xlwings.XlwingsError(
            f"Your SharePoint configuration either requires your workbook name to be "
            f"unique across all synced SharePoint folders or you need to "
            f"{'edit' if sharepoint_config else 'add'} the {sharepoint_config_name} "
            f"setting including one or more folder levels, see: xlwings.org/error."
        )


def to_pdf(
    obj,
    path=None,
    include=None,
    exclude=None,
    layout=None,
    exclude_start_string=None,
    show=None,
    quality=None,
):
    report_path = fspath(path)
    layout_path = fspath(layout)
    if isinstance(obj, (xlwings.Book, xlwings.Sheet)):
        if report_path is None:
            filename, extension = os.path.splitext(obj.fullname)
            directory, _ = os.path.split(obj.fullname)
            if directory:
                report_path = os.path.join(directory, filename + ".pdf")
            else:
                report_path = filename + ".pdf"
        if (include is not None) and (exclude is not None):
            raise ValueError("You can only use either 'include' or 'exclude'")
        # Hide sheets to exclude them from printing
        if isinstance(include, (str, int)):
            include = [include]
        if isinstance(exclude, (str, int)):
            exclude = [exclude]
        exclude_by_name = [
            sheet.index
            for sheet in obj.sheets
            if sheet.name.startswith(exclude_start_string)
        ]
        visibility = {}
        if include or exclude or exclude_by_name:
            for sheet in obj.sheets:
                visibility[sheet] = sheet.visible
        try:
            if include:
                for sheet in obj.sheets:
                    if (sheet.name in include) or (sheet.index in include):
                        sheet.visible = True
                    else:
                        sheet.visible = False
            if exclude or exclude_by_name:
                exclude = [] if exclude is None else exclude
                for sheet in obj.sheets:
                    if (
                        (sheet.name in exclude)
                        or (sheet.index in exclude)
                        or (sheet.index in exclude_by_name)
                    ):
                        sheet.visible = False
            obj.impl.to_pdf(os.path.realpath(report_path), quality=quality)
        except Exception:
            raise
        finally:
            # Reset visibility
            if include or exclude or exclude_by_name:
                for sheet, tf in visibility.items():
                    sheet.visible = tf
    else:
        if report_path is None:
            if isinstance(obj, xlwings.Chart):
                directory, _ = os.path.split(obj.parent.book.fullname)
                filename = obj.name
            elif isinstance(obj, xlwings.Range):
                directory, _ = os.path.split(obj.sheet.book.fullname)
                filename = (
                    str(obj)
                    .replace("<", "")
                    .replace(">", "")
                    .replace(":", "_")
                    .replace(" ", "")
                )
            else:
                raise ValueError(f"Object of type {type(obj)} are not supported.")
            if directory:
                report_path = os.path.join(directory, filename + ".pdf")
            else:
                report_path = filename + ".pdf"
        obj.impl.to_pdf(os.path.realpath(report_path), quality=quality)

    if layout:
        from .pro.reports.pdf import print_on_layout

        print_on_layout(report_path=report_path, layout_path=layout_path)

    if show:
        if sys.platform.startswith("win"):
            os.startfile(report_path)
        else:
            subprocess.run(["open", report_path])
    return report_path
