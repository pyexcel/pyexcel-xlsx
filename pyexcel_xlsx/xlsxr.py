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
from pyexcel_io._compact import OrderedDict, irange


class MergedCell(object):
    def __init__(self, cell_ranges_str):
        print(cell_ranges_str)
        topleft, bottomright = cell_ranges_str.split(':')
        self.__rl, self.__cl = convert_coordinate(topleft)
        self.__rh, self.__ch = convert_coordinate(bottomright)
        self.value = None

    def register_cells(self, registry):
        for rowx in irange(self.__rl, self.__rh+1):
            for colx in irange(self.__cl, self.__ch+1):
                key = "%s-%s" % (rowx, colx)
                registry[key] = self


def convert_coordinate(cell_coordinate_with_letter):
    xy = openpyxl.utils.coordinate_from_string(cell_coordinate_with_letter)
    col = openpyxl.utils.column_index_from_string(xy[0])
    row = xy[1]
    return row, col


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
    def __init__(self, sheet, **keywords):
        SheetReader.__init__(self, sheet, **keywords)
        print(sheet.max_row)
        self.__merged_cells = {}
        for ranges_str in sheet.merged_cell_ranges:
            merged_cells = MergedCell(ranges_str)
            merged_cells.register_cells(self.__merged_cells)

    def row_iterator(self):
        """
        skip hidden rows
        """
        for row_index, row in enumerate(self._native_sheet.iter_rows(), 1):
            if self._native_sheet.row_dimensions[row_index].hidden is False:
                yield (row, row_index)

    def column_iterator(self, row_struct):
        """
        skip hidden columns
        """
        row, row_index = row_struct
        for column_index, cell in enumerate(row, 1):
            letter = openpyxl.utils.get_column_letter(column_index)
            if self._native_sheet.column_dimensions[letter].hidden is False:
                value = cell.value
                if self.__merged_cells:
                    merged_cell = self.__merged_cells.get("%s-%s" % (
                        row_index, column_index))
                    if merged_cell:
                        if merged_cell.value:
                            value = merged_cell.value
                        else:
                            merged_cell.value = value
                yield value


class XLSXBook(BookReader):
    """
    Open xlsx as read only mode
    """
    def open(self, file_name, skip_hidden_sheets=True,
             detect_merged_cells=False,
             skip_hidden_row_and_column=True, **keywords):
        BookReader.open(self, file_name, **keywords)
        self.skip_hidden_sheets = skip_hidden_sheets
        self.skip_hidden_row_and_column = skip_hidden_row_and_column
        self.detect_merged_cells = detect_merged_cells
        self._load_the_excel_file(file_name)

    def open_stream(self, file_stream, skip_hidden_sheets=True,
                    detect_merged_cells=False,
                    skip_hidden_row_and_column=True, **keywords):
        BookReader.open_stream(self, file_stream, **keywords)
        self.skip_hidden_sheets = skip_hidden_sheets
        self.skip_hidden_row_and_column = skip_hidden_row_and_column
        self.detect_merged_cells = detect_merged_cells
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
        if self.skip_hidden_row_and_column or self.detect_merged_cells:
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
