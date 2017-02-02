import os
from unittest import TestCase
from textwrap import dedent

import pyexcel as pe


class TestDateFormat:
    def test_reading_date_format(self):
        """
        date     time
        25/12/14 11:11:11
        25/12/14 12:11:11
        01/01/15 13:13:13
        0.0      0.0
        """
        import datetime
        r = pe.get_sheet(file_name=os.path.join("tests", "fixtures",
                                                "date_field.xlsx"),
                         library='pyexcel-xlsx')
        assert isinstance(r[1, 0], datetime.date) is True
        assert r[1, 0].strftime("%d/%m/%y") == "25/12/14"
        assert isinstance(r[1, 1], datetime.time) is True
        assert r[1, 1].strftime("%H:%M:%S") == "11:11:11"
        value = r[4, 0].isoformat()
        assert value == "1899-12-30T00:00:00"
        assert r[4, 1].isoformat() == "1899-12-30T00:00:00"

    def test_writing_date_format(self):
        import datetime
        excel_filename = "testdateformat.xlsx"
        data = [[datetime.date(2014, 12, 25),
                datetime.time(11, 11, 11),
                datetime.datetime(2014, 12, 25, 11, 11, 11)]]
        pe.save_as(dest_file_name=excel_filename, array=data)
        r = pe.get_sheet(file_name=excel_filename, library='pyexcel-xlsx')
        assert isinstance(r[0, 0], datetime.date) is True
        assert r[0, 0].strftime("%d/%m/%y") == "25/12/14"
        assert isinstance(r[0, 1], datetime.time) is True
        assert r[0, 1].strftime("%H:%M:%S") == "11:11:11"
        assert isinstance(r[0, 2], datetime.date) is True
        assert r[0, 2].strftime("%d/%m/%y %H:%M:%S") == "25/12/14 11:11:11"
        os.unlink(excel_filename)


class TestAutoDetectInt(TestCase):
    def setUp(self):
        self.content = [[1, 2, 3.1]]
        self.test_file = "test_auto_detect_init.xlsx"
        pe.save_as(array=self.content, dest_file_name=self.test_file)

    def test_auto_detect_int(self):
        sheet = pe.get_sheet(file_name=self.test_file)
        expected = dedent("""
        pyexcel_sheet1:
        +---+---+-----+
        | 1 | 2 | 3.1 |
        +---+---+-----+""").strip()
        self.assertEqual(str(sheet), expected)

    def test_get_book_auto_detect_int(self):
        book = pe.get_book(file_name=self.test_file)
        expected = dedent("""
        pyexcel_sheet1:
        +---+---+-----+
        | 1 | 2 | 3.1 |
        +---+---+-----+""").strip()
        self.assertEqual(str(book), expected)

    def test_auto_detect_int_false(self):
        sheet = pe.get_sheet(file_name=self.test_file, auto_detect_int=False)
        expected = dedent("""
        pyexcel_sheet1:
        +-----+-----+-----+
        | 1.0 | 2.0 | 3.1 |
        +-----+-----+-----+""").strip()
        self.assertEqual(str(sheet), expected)

    def test_get_book_auto_detect_int_false(self):
        book = pe.get_book(file_name=self.test_file, auto_detect_int=False)
        expected = dedent("""
        pyexcel_sheet1:
        +-----+-----+-----+
        | 1.0 | 2.0 | 3.1 |
        +-----+-----+-----+""").strip()
        self.assertEqual(str(book), expected)

    def tearDown(self):
        os.unlink(self.test_file)
