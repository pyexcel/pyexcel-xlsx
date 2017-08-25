"""
    pyexcel_xlsx.xlsxr
    ~~~~~~~~~~~~~~~~~~~

    Read xlsx file format using openpyxl

    :copyright: (c) 2015-2017 by Onni Software Ltd & its contributors
    :license: New BSD License
"""
import openpyxl

from pyexcel_io.book import BookReader
from pyexcel_io.sheet import SheetReader
from pyexcel_io._compact import OrderedDict


class XLSXSheet(SheetReader):
    """
    Iterate through rows
    """
    @property
    def name(self):
        """sheet name"""
        return self._native_sheet.title

    def row_iterator(self):
        """
        openpyxl row iterator

        http://openpyxl.readthedocs.io/en/default/optimized.html
        """
        return self._native_sheet.rows

    def column_iterator(self, row):
        """
        a generator for the values in a row
        """
        for cell in row:
            yield cell.value


class XLSXBook(BookReader):
    """
    Open xlsx as read only mode
    """
    def open(self, file_name, skip_hidden_sheets=True, **keywords):
        BookReader.open(self, file_name, **keywords)
        self.skip_hidden_sheets = skip_hidden_sheets
        self._load_the_excel_file(file_name)

    def open_stream(self, file_stream, skip_hidden_sheets=True, **keywords):
        BookReader.open_stream(self, file_stream, **keywords)
        self.skip_hidden_sheets = skip_hidden_sheets
        self._load_the_excel_file(file_stream)

    def read_sheet_by_name(self, sheet_name):
        sheet = self._native_book.get_sheet_by_name(sheet_name)
        if sheet is None:
            raise ValueError("%s cannot be found" % sheet_name)
        else:
            return self.read_sheet(sheet)

    def read_sheet_by_index(self, sheet_index):
        names = self._native_book.sheetnames
        length = len(names)
        if sheet_index < length:
            return self.read_sheet_by_name(names[sheet_index])
        else:
            raise IndexError("Index %d of out bound %d" % (
                sheet_index,
                length))

    def read_all(self):
        result = OrderedDict()
        for sheet in self._native_book:
            if self.skip_hidden_sheets and sheet.sheet_state == 'hidden':
                continue
            data_dict = self.read_sheet(sheet)
            result.update(data_dict)
        return result

    def read_sheet(self, native_sheet):
        sheet = XLSXSheet(native_sheet, **self._keywords)
        return {sheet.name: sheet.to_array()}

    def close(self):
        self._native_book.close()
        self._native_book = None

    def _load_the_excel_file(self, file_alike_object):
        self._native_book = openpyxl.load_workbook(
            filename=file_alike_object, data_only=True, read_only=True)
