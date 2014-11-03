def int_to_rgb(number):
    """Given a int sequel, return the rgb"""
    number = int(number)
    r = number % 256
    g = (number / 256) % 256
    b = (number / (256 * 256)) % 256
    return r, g, b


def rgb_to_int(rgb):
    """Given an rgb, return an int"""
    return rgb[0] + (rgb[1] * 256) + (rgb[2] * 256 * 256)