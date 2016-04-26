"""
    pyexcel_xlsx
    ~~~~~~~~~~~~~~~~~~~

    The lower level xlsx file format handler using openpyxl

    :copyright: (c) 2015-2016 by Onni Software Ltd & its contributors
    :license: New BSD License
"""
from pyexcel_io.io import get_data as read_data, isstream, store_data as write_data


def save_data(afile, data, file_type=None, **keywords):
    if isstream(afile) and file_type is None:
        file_type='xlsx'
    write_data(afile, data, file_type=file_type, **keywords)


def get_data(afile, file_type=None, **keywords):
    if isstream(afile) and file_type is None:
        file_type='xlsx'
    return read_data(afile, file_type=file_type, **keywords)

