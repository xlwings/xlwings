from ...utils import fspath


class Image:
    """
    Use this class to provide images to either ``render_template()``.

    Arguments
    ---------

    filename : str or pathlib.Path object
        The file name or path
    """
    def __init__(self, filename):
        self.filename = fspath(filename)
