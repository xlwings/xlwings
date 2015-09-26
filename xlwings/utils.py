from __future__ import division
import datetime as dt

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