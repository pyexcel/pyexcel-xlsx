import os

import pyexcel
from base import create_sample_file1

from nose.tools import eq_


class TestStringIO:
    def test_xlsx_stringio(self):
        testfile = "cute.xlsx"
        create_sample_file1(testfile)
        with open(testfile, "rb") as f:
            content = f.read()
            r = pyexcel.get_sheet(
                file_type="xlsx", file_content=content, library="pyexcel-xlsx"
            )
            result = ["a", "b", "c", "d", "e", "f", "g", "h", "i", "j", 1.1, 1]
            actual = list(r.enumerate())
            eq_(result, actual)
        if os.path.exists(testfile):
            os.unlink(testfile)

    def test_xlsx_output_stringio(self):
        data = [[1, 2, 3], [4, 5, 6]]
        io = pyexcel.save_as(dest_file_type="xlsx", array=data)
        r = pyexcel.get_sheet(
            file_type="xlsx",
            file_content=io.getvalue(),
            library="pyexcel-xlsx",
        )
        result = [1, 2, 3, 4, 5, 6]
        actual = list(r.enumerate())
        eq_(result, actual)
