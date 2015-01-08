"""

  This file keeps all fixes for issues found

"""

import os
import pyexcel as pe
import pyexcel.ext.xls
import pyexcel.ext.xlsx
import datetime
from pyexcel.ext.xlsx import get_columns
from pyexcel.sheets.matrix import _excel_column_index


class TestBugFix:
    def test_pyexcel_issue_4(self):
        """pyexcel issue #4"""
        indices = [
            'A',
            'AA',
            'ABC',
            'ABCD',
            'ABCDE',
            'ABCDEF',
            'ABCDEFG',
            'ABCDEFGH',
            'ABCDEFGHI',
            'ABCDEFGHIJ',
            'ABCDEFGHIJK',
            'ABCDEFGHIJKL',
            'ABCDEFGHIJKLM',
            'ABCDEFGHIJKLMN',
            'ABCDEFGHIJKLMNO',
            'ABCDEFGHIJKLMNOP',
            'ABCDEFGHIJKLMNOPQ',
            'ABCDEFGHIJKLMNOPQR',
            'ABCDEFGHIJKLMNOPQRS',
            'ABCDEFGHIJKLMNOPQRST',
            'ABCDEFGHIJKLMNOPQRSTU',
            'ABCDEFGHIJKLMNOPQRSTUV',
            'ABCDEFGHIJKLMNOPQRSTUVW',
            'ABCDEFGHIJKLMNOPQRSTUVWX',
            'ABCDEFGHIJKLMNOPQRSTUVWXY',
            'ABCDEFGHIJKLMNOPQRSTUVWXYZ',
        ]
        for column_name in indices:
            print("Testing %s" % column_name)
            column_index = _excel_column_index(column_name)
            new_column_name = get_columns(column_index)
            assert new_column_name == column_name

    def test_pyexcel_issue_5(self):
        """pyexcel issue #5

        datetime is not properly parsed
        """
        s = pe.load(os.path.join("tests",
                                 "test-fixtures",
                                 "test-date-format.xls"))
        s.save_as("issue5.xlsx")
        s2 = pe.load("issue5.xlsx")
        assert s[0,0] == datetime.datetime(2015, 11, 11, 11, 12, 0)
        assert s2[0,0] == datetime.datetime(2015, 11, 11, 11, 12, 0)
