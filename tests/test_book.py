import os
from datetime import time, datetime

from pyexcel_xlsx import get_data
from pyexcel_xlsx.book import Book
from pyexcel_io._compact import OrderedDict

from nose.tools import eq_


def test_book():
    test_file = "book_test.xlsx"
    book = Book(os.path.join("tests", "fixtures", "date_field.xlsx"))
    sheet = book["Sheet2"]
    sheet["A1"] = "Test"
    book.save(test_file)
    book.close()

    data = get_data(test_file)
    eq_(data["Sheet2"], [["Test"]])
    os.unlink(test_file)


def test_create_new_sheet():
    test_file = "book_test.xlsx"
    book = Book(os.path.join("tests", "fixtures", "date_field.xlsx"))
    sheet = book["alien"]
    sheet["A1"] = "Test"
    book.save(test_file)
    book.close()

    data = get_data(test_file)
    eq_(data["alien"], [["Test"]])
    os.unlink(test_file)


def test_create_append_new_data():
    test_file = "book_test.xlsx"
    book = Book(os.path.join("tests", "fixtures", "date_field.xlsx"))
    sheet = book["alien"]
    sheet.append([1, 2, 3])
    sheet.append([3, 4, 5])
    book.save(test_file)
    book.close()

    data = get_data(test_file)
    eq_(data["alien"], [[1, 2, 3], [3, 4, 5]])
    os.unlink(test_file)
