"""
    pyexcel_xlsx
    ~~~~~~~~~~~~~~~~~~~

    The lower level xlsx file format handler using openpyxl

    :copyright: (c) 2015-2016 by Onni Software Ltd & its contributors
    :license: New BSD License
"""
# flake8: noqa
# this line has to be place above all else
# because of dynamic import
__FILE_TYPE__ = 'xlsx'
__META__ = {
    'submodule': __FILE_TYPE__,
    'file_types': [__FILE_TYPE__, 'xlsm'],
    'stream_type': 'binary'
}

__pyexcel_io_plugins__ = [__META__]


from pyexcel_io.io import get_data as read_data, isstream, store_data as write_data


def save_data(afile, data, file_type=None, **keywords):
    """standalone module function for writing module supported file type"""
    if isstream(afile) and file_type is None:
        file_type = __FILE_TYPE__
    write_data(afile, data, file_type=file_type, **keywords)


def get_data(afile, file_type=None, **keywords):
    """standalone module function for reading module supported file type"""
    if isstream(afile) and file_type is None:
        file_type = __FILE_TYPE__
    return read_data(afile, file_type=file_type, **keywords)
