"""

  This file keeps all fixes for issues found

"""
import os
import sys
import datetime
from textwrap import dedent

import pyexcel as pe

from nose.tools import eq_

IN_TRAVIS = "TRAVIS" in os.environ


PY36_ABOVE = sys.version_info[0] == 3 and sys.version_info[1] >= 6


def test_pyexcel_issue_5():
    """pyexcel issue #5

    datetime is not properly parsed
    """
    s = pe.get_sheet(file_name=get_fixtures("test-date-format.xls"))
    s.save_as("issue5.xlsx")
    s2 = pe.load("issue5.xlsx")
    assert s[0, 0] == datetime.datetime(2015, 11, 11, 11, 12, 0)
    assert s2[0, 0] == datetime.datetime(2015, 11, 11, 11, 12, 0)


def test_pyexcel_issue_8_with_physical_file():
    """pyexcel issue #8

    formular got lost
    """
    tmp_file = "issue_8_save_as.xlsx"
    s = pe.get_sheet(file_name=get_fixtures("test8.xlsx"))
    s.save_as(tmp_file)
    s2 = pe.load(tmp_file)
    eq_(str(s), str(s2))
    content = dedent(
        """
    CNY:
    +----------+----------+------+---+-------+
    | 01/09/13 | 02/09/13 | 1000 | 5 | 13.89 |
    +----------+----------+------+---+-------+
    | 02/09/13 | 03/09/13 | 2000 | 6 | 33.33 |
    +----------+----------+------+---+-------+
    | 03/09/13 | 04/09/13 | 3000 | 7 | 58.33 |
    +----------+----------+------+---+-------+"""
    ).strip("\n")
    eq_(str(s2), content)
    os.unlink(tmp_file)


def test_pyexcel_issue_8_with_memory_file():
    """pyexcel issue #8

    formular got lost
    """
    tmp_file = "issue_8_save_as.xlsx"
    f = open(get_fixtures("test8.xlsx"), "rb")
    s = pe.load_from_memory("xlsx", f.read())
    s.save_as(tmp_file)
    s2 = pe.load(tmp_file)
    eq_(str(s), str(s2))
    content = dedent(
        """
    CNY:
    +----------+----------+------+---+-------+
    | 01/09/13 | 02/09/13 | 1000 | 5 | 13.89 |
    +----------+----------+------+---+-------+
    | 02/09/13 | 03/09/13 | 2000 | 6 | 33.33 |
    +----------+----------+------+---+-------+
    | 03/09/13 | 04/09/13 | 3000 | 7 | 58.33 |
    +----------+----------+------+---+-------+"""
    ).strip("\n")
    eq_(str(s2), content)
    os.unlink(tmp_file)


def test_excessive_columns():
    tmp_file = "date_field.xlsx"
    s = pe.get_sheet(file_name=get_fixtures(tmp_file))
    assert s.number_of_columns() == 2


def test_issue_8_hidden_sheet():
    test_file = get_fixtures("hidden_sheets.xlsx")
    book_dict = pe.get_book_dict(file_name=test_file, library="pyexcel-xlsx")
    assert "hidden" not in book_dict
    eq_(book_dict["shown"], [["A", "B"]])


def test_issue_8_hidden_sheet_2():
    test_file = get_fixtures("hidden_sheets.xlsx")
    book_dict = pe.get_book_dict(
        file_name=test_file, skip_hidden_sheets=False, library="pyexcel-xlsx"
    )
    assert "hidden" in book_dict
    eq_(book_dict["shown"], [["A", "B"]])
    eq_(book_dict["hidden"], [["a", "b"]])


def test_issue_20():
    pe.get_book(
        url="https://github.com/pyexcel/pyexcel-xlsx/raw/master/tests/fixtures/file_with_an_empty_sheet.xlsx"  # noqa: E501
    )


def get_fixtures(file_name):
    return os.path.join("tests", "fixtures", file_name)
