class XlwingsError(Exception):
    """Base exception for all custom xlwings errors."""
    pass


class ExcelCallError(XlwingsError):
    """Exception for errors during calling of excel app from python."""
    pass


class PythonCallError(XlwingsError):
    """Exception for errors during calling of python from excel app."""
    pass


class ShapeAlreadyExists(XlwingsError):
    """Exception for errors regarding duplicate excel shapes."""
    pass
