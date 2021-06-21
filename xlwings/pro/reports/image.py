from ...utils import fspath


class Image:
    """
    filename : str or pathlib.Path object
        The file name or path
    """
    def __init__(self, filename):
        self.filename = fspath(filename)
