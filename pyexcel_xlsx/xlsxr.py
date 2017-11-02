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
        for row in self._native_sheet.rows:
            yield row

    def column_iterator(self, row):
        """
        a generator for the values in a row
        """
        for cell in row:
            yield cell.value


class SlowSheet(XLSXSheet):
    """
    This sheet will be slower because it does not use readonly sheet
    """
    def row_iterator(self):
        """
        skip hidden rows
        """
        for row_index, row in enumerate(self._native_sheet.rows, 1):
            if self._native_sheet.row_dimensions[row_index].hidden is False:
                yield row

    def column_iterator(self, row):
        """
        skip hidden columns
        """
        for column_index, cell in enumerate(row, 1):
            letter = openpyxl.utils.get_column_letter(column_index)
            if self._native_sheet.column_dimensions[letter].hidden is False:
                yield cell.value


class XLSXBook(BookReader):
    """
    Open xlsx as read only mode
    """
    def open(self, file_name, skip_hidden_sheets=True,
             skip_hidden_row_and_column=True, **keywords):
        BookReader.open(self, file_name, **keywords)
        self.skip_hidden_sheets = skip_hidden_sheets
        self.skip_hidden_row_and_column = skip_hidden_row_and_column
        self._load_the_excel_file(file_name)

    def open_stream(self, file_stream, skip_hidden_sheets=True,
                    skip_hidden_row_and_column=True, **keywords):
        BookReader.open_stream(self, file_stream, **keywords)
        self.skip_hidden_sheets = skip_hidden_sheets
        self.skip_hidden_row_and_column = skip_hidden_row_and_column
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
        if self.skip_hidden_row_and_column:
            sheet = SlowSheet(native_sheet, **self._keywords)
        else:
            sheet = XLSXSheet(native_sheet, **self._keywords)
        return {sheet.name: sheet.to_array()}

    def close(self):
        self._native_book.close()
        self._native_book = None

    def _load_the_excel_file(self, file_alike_object):
        read_only_flag = True
        if self.skip_hidden_row_and_column:
            read_only_flag = False
        self._native_book = openpyxl.load_workbook(
            filename=file_alike_object, data_only=True,
            read_only=read_only_flag)
