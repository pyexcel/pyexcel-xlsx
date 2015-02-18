from base import PyexcelMultipleSheetBase
import pyexcel
import os
from pyexcel.ext import xlsx
import sys

if sys.version_info[0] == 2 and sys.version_info[1] < 7:
    from ordereddict import OrderedDict
else:
    from collections import OrderedDict


class TestXlsmNxlsMultipleSheets(PyexcelMultipleSheetBase):
    def setUp(self):
        self.testfile = "multiple1.xlsm"
        self.testfile2 = "multiple1.xlsx"
        self.content = {
            "Sheet1": [[1, 1, 1, 1], [2, 2, 2, 2], [3, 3, 3, 3]],
            "Sheet2": [[4, 4, 4, 4], [5, 5, 5, 5], [6, 6, 6, 6]],
            "Sheet3": [[u'X', u'Y', u'Z'], [1, 4, 7], [2, 5, 8], [3, 6, 9]]
        }
        self._write_test_file(self.testfile)

    def tearDown(self):
        self._clean_up()


class TestXlsNXlsxMultipleSheets(PyexcelMultipleSheetBase):
    def setUp(self):
        self.testfile = "multiple1.xlsm"
        self.testfile2 = "multiple1.xlsx"
        self.content = OrderedDict()
        self.content.update({"Sheet1": [[1, 1, 1, 1], [2, 2, 2, 2], [3, 3, 3, 3]]})
        self.content.update({"Sheet2": [[4, 4, 4, 4], [5, 5, 5, 5], [6, 6, 6, 6]]})
        self.content.update({"Sheet3": [[u'X', u'Y', u'Z'], [1, 4, 7], [2, 5, 8], [3, 6, 9]]})
        self._write_test_file(self.testfile)

    def tearDown(self):
        self._clean_up()


class TestAddBooks:
    def _write_test_file(self, file):
        """
        Make a test file as:

        1,1,1,1
        2,2,2,2
        3,3,3,3
        """
        self.rows = 3
        w = pyexcel.BookWriter(file)
        w.write_book_from_dict(self.content)
        w.close()

    def setUp(self):
        self.testfile = "multiple3.xlsx"
        self.testfile2 = "multiple1.xlsx"
        self.testfile3 = "multiple2.xlsx"
        self.content = {
            "Sheet1": [[1, 1, 1, 1], [2, 2, 2, 2], [3, 3, 3, 3]],
            "Sheet2": [[4, 4, 4, 4], [5, 5, 5, 5], [6, 6, 6, 6]],
            "Sheet3": [[u'X', u'Y', u'Z'], [1, 4, 7], [2, 5, 8], [3, 6, 9]]
        }
        self._write_test_file(self.testfile)
        self._write_test_file(self.testfile2)

    def test_load_a_single_sheet(self):
        b1 = pyexcel.load_book(self.testfile, sheet_name="Sheet1")
        assert len(b1.sheet_names()) == 1
        assert b1['Sheet1'].to_array() == self.content['Sheet1']

    def test_load_a_single_sheet2(self):
        b1 = pyexcel.load_book(self.testfile, sheet_index=0)
        assert len(b1.sheet_names()) == 1
        assert b1['Sheet1'].to_array() == self.content['Sheet1']

    def test_delete_sheets(self):
        b1 = pyexcel.load_book(self.testfile)
        assert len(b1.sheet_names()) == 3
        del b1["Sheet1"]
        assert len(b1.sheet_names()) == 2
        try:
            del b1["Sheet1"]
            assert 1==2
        except KeyError:
            assert 1==1
        del b1[1]
        assert len(b1.sheet_names()) == 1
        try:
            del b1[1]
            assert 1==2
        except IndexError:
            assert 1==1
            
    def test_delete_sheets2(self):
        """repetitively delete first sheet"""
        b1 = pyexcel.load_book(self.testfile)
        del b1[0]
        assert len(b1.sheet_names()) == 2
        del b1[0]
        assert len(b1.sheet_names()) == 1
        del b1[0]
        assert len(b1.sheet_names()) == 0
        
    def test_add_book1(self):
        """
        test this scenario: book3 = book1 + book2
        """
        b1 = pyexcel.BookReader(self.testfile)
        b2 = pyexcel.BookReader(self.testfile2)
        b3 = b1 + b2
        content = pyexcel.utils.to_dict(b3)
        sheet_names = content.keys()
        assert len(sheet_names) == 6
        for name in sheet_names:
            if "Sheet3" in name:
                assert content[name] == self.content["Sheet3"]
            elif "Sheet2" in name:
                assert content[name] == self.content["Sheet2"]
            elif "Sheet1" in name:
                assert content[name] == self.content["Sheet1"]
        
    def test_add_book1_in_place(self):
        """
        test this scenario book1 +=  book2
        """
        b1 = pyexcel.BookReader(self.testfile)
        b2 = pyexcel.BookReader(self.testfile2)
        b1 += b2
        content = pyexcel.utils.to_dict(b1)
        sheet_names = content.keys()
        assert len(sheet_names) == 6
        for name in sheet_names:
            if "Sheet3" in name:
                assert content[name] == self.content["Sheet3"]
            elif "Sheet2" in name:
                assert content[name] == self.content["Sheet2"]
            elif "Sheet1" in name:
                assert content[name] == self.content["Sheet1"]

    def test_add_book2(self):
        """
        test this scenario book3 = book1 + sheet3
        """
        b1 = pyexcel.BookReader(self.testfile)
        b2 = pyexcel.BookReader(self.testfile2)
        b3 = b1 + b2["Sheet3"]
        content = pyexcel.utils.to_dict(b3)
        sheet_names = content.keys()
        assert len(sheet_names) == 4
        for name in sheet_names:
            if "Sheet3" in name:
                assert content[name] == self.content["Sheet3"]
            elif "Sheet2" in name:
                assert content[name] == self.content["Sheet2"]
            elif "Sheet1" in name:
                assert content[name] == self.content["Sheet1"]

    def test_add_book2_in_place(self):
        """
        test this scenario book3 = book1 + sheet3
        """
        b1 = pyexcel.BookReader(self.testfile)
        b2 = pyexcel.BookReader(self.testfile2)
        b1 += b2["Sheet3"]
        content = pyexcel.utils.to_dict(b1)
        sheet_names = content.keys()
        assert len(sheet_names) == 4
        for name in sheet_names:
            if "Sheet3" in name:
                assert content[name] == self.content["Sheet3"]
            elif "Sheet2" in name:
                assert content[name] == self.content["Sheet2"]
            elif "Sheet1" in name:
                assert content[name] == self.content["Sheet1"]

    def test_add_book3(self):
        """
        test this scenario book3 = sheet1 + sheet2
        """
        b1 = pyexcel.BookReader(self.testfile)
        b2 = pyexcel.BookReader(self.testfile2)
        b3 = b1["Sheet1"] + b2["Sheet3"]
        content = pyexcel.utils.to_dict(b3)
        sheet_names = content.keys()
        assert len(sheet_names) == 2
        assert content["Sheet3"] == self.content["Sheet3"]
        assert content["Sheet1"] == self.content["Sheet1"]
        
    def test_add_book4(self):
        """
        test this scenario book3 = sheet1 + book
        """
        b1 = pyexcel.BookReader(self.testfile)
        b2 = pyexcel.BookReader(self.testfile2)
        b3 = b1["Sheet1"] + b2
        content = pyexcel.utils.to_dict(b3)
        sheet_names = content.keys()
        assert len(sheet_names) == 4
        for name in sheet_names:
            if "Sheet3" in name:
                assert content[name] == self.content["Sheet3"]
            elif "Sheet2" in name:
                assert content[name] == self.content["Sheet2"]
            elif "Sheet1" in name:
                assert content[name] == self.content["Sheet1"]

    def test_add_book_error(self):
        """
        test this scenario: book3 = sheet1 + book
        """
        b1 = pyexcel.BookReader(self.testfile)
        try:
            b1 + 12
            assert 1==2
        except TypeError:
            assert 1==1
        try:
            b1 += 12
            assert 1==2
        except TypeError:
            assert 1==1

    def tearDown(self):
        if os.path.exists(self.testfile):
            os.unlink(self.testfile)
        if os.path.exists(self.testfile2):
            os.unlink(self.testfile2)


class TestMultiSheetReader:
    def setUp(self):
        self.testfile = "file_with_an_empty_sheet.xlsx"

    def test_reader_with_correct_sheets(self):
        r = pyexcel.BookReader(os.path.join("tests", "fixtures", self.testfile))
        assert r.number_of_sheets() == 3
