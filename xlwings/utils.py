from __future__ import division


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