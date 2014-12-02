from .xlsxbook import XLSXBook, XLSXWriter
try:
    from pyexcel.io import READERS
    from pyexcel.io import WRITERS

    READERS.update({
        "xlsm": XLSXBook,
        "xlsx": XLSXBook
    })
    WRITERS.update({
        "xlsm": XLSXWriter,
        "xlsx": XLSXWriter
    })
except:
    # to allow this module to function independently
    pass

__VERSION__ = "0.0.1"