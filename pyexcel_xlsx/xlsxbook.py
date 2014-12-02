"""
    pyexcel.ext.xlbook
    ~~~~~~~~~~~~~~~~~~~

    The lower level xls/xlsx/xlsm file format handler using xlrd/xlwt

    :copyright: (c) 2014 by C. W.
    :license: GPL v3
"""
import sys
import openpyxl
from pyexcel.ioext import SheetReader, BookReader, SheetWriter, BookWriter
if sys.version_info[0] < 3:
    from StringIO import StringIO
else:
    from io import BytesIO as StringIO

COLUMNS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ'
COLUMN_LENGTH = 26

def get_columns(index):
    if index < COLUMN_LENGTH:
        return COLUMNS[index]
    else:
        return get_columns(int(index/COLUMN_LENGTH)) + COLUMNS[index%COLUMN_LENGTH]
    

class XLSXSheet(SheetReader):
    """
    xls sheet

    Currently only support first sheet in the file
    """
    def number_of_rows(self):
        """
        Number of rows in the xls sheet
        """
        return self.worksheet.get_highest_row()

    def number_of_columns(self):
        """
        Number of columns in the xls sheet
        """
        return self.worksheet.get_highest_column()

    def cell_value(self, row, column):
        """
        Random access to the xls cells
        """
        actual_row = row + 1
        return self.worksheet.cell("%s%d" % (get_columns(column), actual_row)).value

class XLSXBook(BookReader):
    """
    XLSBook reader

    It reads xls, xlsm, xlsx work book
    """
    def getSheet(self, nativeSheet):
        return XLSXSheet(nativeSheet)

    def load_from_memory(self, file_content):
        return openpyxl.load_workbook(filename=StringIO(file_content))

    def load_from_file(self, filename):
        return openpyxl.load_workbook(filename)

    def sheetIterator(self):
        return self.workbook


class XLSXSheetWriter(SheetWriter):
    """
    xls, xlsx and xlsm sheet writer
    """
    def set_sheet_name(self, name):
        self.sheet.title = name
        self.current_row = 1

    def write_row(self, array):
        """
        write a row into the file
        """
        for i in range(0, len(array)):
            self.sheet.cell("%s%d" % (get_columns(i), self.current_row)).value = array[i]
        self.current_row += 1


class XLSXWriter(BookWriter):
    """
    xls, xlsx and xlsm writer
    """
    def __init__(self, file):
        BookWriter.__init__(self, file)
        self.wb = openpyxl.Workbook()
        self.current_sheet = 0

    def create_sheet(self, name):
        if self.current_sheet == 0:
            self.current_sheet = 1
            return XLSXSheetWriter(self.wb.active, name)
        else:
            
            return XLSXSheetWriter(self.wb.create_sheet(), name)

    def close(self):
        """
        This call actually save the file
        """
        self.wb.save(filename=self.file)
