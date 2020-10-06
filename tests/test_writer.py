import os
from collections import OrderedDict

from base import PyexcelWriterBase, PyexcelHatWriterBase
from pyexcel_io.reader import EncapsulatedSheetReader
from pyexcel_xlsx.xlsxr import XLSXBook as Reader
from pyexcel_xlsx.xlsxw import XLSXWriter as Writer


class TestNativeXLSXWriter:
    def test_write_book(self):
        self.content = {
            "Sheet1": [[1, 1, 1, 1], [2, 2, 2, 2], [3, 3, 3, 3]],
            "Sheet2": [[4, 4, 4, 4], [5, 5, 5, 5], [6, 6, 6, 6]],
            "Sheet3": [[u"X", u"Y", u"Z"], [1, 4, 7], [2, 5, 8], [3, 6, 9]],
        }
        self.testfile = "writer.xlsx"
        writer = Writer(self.testfile, "xlsx")
        writer.write(self.content)
        writer.close()
        reader = Reader(self.testfile, "xlsx")
        content = read_all(reader)
        for key in content.keys():
            content[key] = list(content[key])
        assert content == self.content
        reader.close()

    def tearDown(self):
        if os.path.exists(self.testfile):
            os.unlink(self.testfile)


class TestxlsxnCSVWriter(PyexcelWriterBase):
    def setUp(self):
        self.testfile = "test.xlsx"
        self.testfile2 = "test.csv"

    def tearDown(self):
        if os.path.exists(self.testfile):
            os.unlink(self.testfile)
        if os.path.exists(self.testfile2):
            os.unlink(self.testfile2)


class TestxlsxHatWriter(PyexcelHatWriterBase):
    def setUp(self):
        self.testfile = "test.xlsx"

    def tearDown(self):
        if os.path.exists(self.testfile):
            os.unlink(self.testfile)


def read_all(reader):
    result = OrderedDict()
    for index, sheet in enumerate(reader.content_array):
        result.update(
            {
                sheet.name: EncapsulatedSheetReader(
                    reader.read_sheet(index)
                ).to_array()
            }
        )
    return result
