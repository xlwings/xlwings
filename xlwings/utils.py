from __future__ import division
import datetime as dt


missing = object()


try:
    import numpy as np
except ImportError:
    np = None


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


class VBAWriter(object):

    class Block(object):
        def __init__(self, writer, start):
            self.writer = writer
            self.start = start

        def __enter__(self):
            self.writer.writeln(self.start)
            self.writer._indent += 1

        def __exit__(self, exc_type, exc_val, exc_tb):
            self.writer._indent -= 1
            #self.writer.writeln(self.end)

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
        if self._freshline:
            self.f.write('\t' * self._indent)
            self._freshline = False
        if kwargs:
            template = template.format(**kwargs)
        self.f.write(template)
        if template[-1] == '\n':
            self._freshline = True

    def write_label(self, label):
        self._indent -= 1
        self.write(label + ':\n')
        self._indent += 1

    def writeln(self, template, **kwargs):
        self.write(template + '\n', **kwargs)
