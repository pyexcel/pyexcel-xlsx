import os
from collections import OrderedDict

import pyexcel
from base import PyexcelMultipleSheetBase

from nose.tools import raises


class TestXlsmNxlsMultipleSheets(PyexcelMultipleSheetBase):
    def setUp(self):
        self.testfile = "multiple1.xlsm"
        self.testfile2 = "multiple1.xlsx"
        self.content = _produce_ordered_dict()
        self._write_test_file(self.testfile)

    def tearDown(self):
        self._clean_up()


class TestXlsNXlsxMultipleSheets(PyexcelMultipleSheetBase):
    def setUp(self):
        self.testfile = "multiple1.xlsm"
        self.testfile2 = "multiple1.xlsx"
        self.content = _produce_ordered_dict()
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
        pyexcel.save_book_as(bookdict=self.content, dest_file_name=file)

    def setUp(self):
        self.testfile = "multiple3.xlsx"
        self.testfile2 = "multiple1.xlsx"
        self.testfile3 = "multiple2.xlsx"
        self.content = _produce_ordered_dict()
        self._write_test_file(self.testfile)
        self._write_test_file(self.testfile2)

    def test_load_a_single_sheet(self):
        b1 = pyexcel.get_book(
            file_name=self.testfile,
            sheet_name="Sheet1",
            library="pyexcel-xlsx",
        )
        assert len(b1.sheet_names()) == 1
        assert b1["Sheet1"].to_array() == self.content["Sheet1"]

    def test_load_a_single_sheet2(self):
        b1 = pyexcel.get_book(
            file_name=self.testfile, sheet_index=1, library="pyexcel-xlsx"
        )
        assert len(b1.sheet_names()) == 1
        assert b1["Sheet2"].to_array() == self.content["Sheet2"]

    @raises(IndexError)
    def test_load_a_single_sheet3(self):
        pyexcel.get_book(file_name=self.testfile, sheet_index=10000)

    @raises(ValueError)
    def test_load_a_single_sheet4(self):
        pyexcel.get_book(
            file_name=self.testfile,
            sheet_name="Not exist",
            library="pyexcel-xlsx",
        )

    def test_delete_sheets(self):
        b1 = pyexcel.load_book(self.testfile)
        assert len(b1.sheet_names()) == 3
        del b1["Sheet1"]
        assert len(b1.sheet_names()) == 2
        try:
            del b1["Sheet1"]
            assert 1 == 2
        except KeyError:
            assert 1 == 1
        del b1[1]
        assert len(b1.sheet_names()) == 1
        try:
            del b1[1]
            assert 1 == 2
        except IndexError:
            assert 1 == 1

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
        b1 = pyexcel.get_book(file_name=self.testfile)
        b2 = pyexcel.get_book(file_name=self.testfile2)
        b3 = b1 + b2
        content = b3.dict
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
        test this scenario: book1 +=  book2
        """
        b1 = pyexcel.BookReader(self.testfile)
        b2 = pyexcel.BookReader(self.testfile2)
        b1 += b2
        content = b1.dict
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
        test this scenario: book3 = book1 + sheet3
        """
        b1 = pyexcel.BookReader(self.testfile)
        b2 = pyexcel.BookReader(self.testfile2)
        b3 = b1 + b2["Sheet3"]
        content = b3.dict
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
        test this scenario: book3 = book1 + sheet3
        """
        b1 = pyexcel.BookReader(self.testfile)
        b2 = pyexcel.BookReader(self.testfile2)
        b1 += b2["Sheet3"]
        content = b1.dict
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
        test this scenario: book3 = sheet1 + sheet2
        """
        b1 = pyexcel.BookReader(self.testfile)
        b2 = pyexcel.BookReader(self.testfile2)
        b3 = b1["Sheet1"] + b2["Sheet3"]
        content = b3.dict
        sheet_names = content.keys()
        assert len(sheet_names) == 2
        assert content["Sheet3"] == self.content["Sheet3"]
        assert content["Sheet1"] == self.content["Sheet1"]

    def test_add_book4(self):
        """
        test this scenario: book3 = sheet1 + book
        """
        b1 = pyexcel.BookReader(self.testfile)
        b2 = pyexcel.BookReader(self.testfile2)
        b3 = b1["Sheet1"] + b2
        content = b3.dict
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
            assert 1 == 2
        except TypeError:
            assert 1 == 1
        try:
            b1 += 12
            assert 1 == 2
        except TypeError:
            assert 1 == 1

    def tearDown(self):
        if os.path.exists(self.testfile):
            os.unlink(self.testfile)
        if os.path.exists(self.testfile2):
            os.unlink(self.testfile2)


class TestMultiSheetReader:
    def setUp(self):
        self.testfile = "file_with_an_empty_sheet.xlsx"

    def test_reader_with_correct_sheets(self):
        r = pyexcel.BookReader(
            os.path.join("tests", "fixtures", self.testfile)
        )
        assert r.number_of_sheets() == 3


def _produce_ordered_dict():
    data_dict = OrderedDict()
    data_dict.update({"Sheet1": [[1, 1, 1, 1], [2, 2, 2, 2], [3, 3, 3, 3]]})
    data_dict.update({"Sheet2": [[4, 4, 4, 4], [5, 5, 5, 5], [6, 6, 6, 6]]})
    data_dict.update(
        {"Sheet3": [[u"X", u"Y", u"Z"], [1, 4, 7], [2, 5, 8], [3, 6, 9]]}
    )
    return data_dict
