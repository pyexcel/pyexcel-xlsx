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