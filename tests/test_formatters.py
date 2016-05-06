import pyexcel as pe
import datetime
import os


class TestDateFormat:
    def test_reading_date_format(self):
        """
        date     time
        25/12/14 11:11:11
        25/12/14 12:11:11
        01/01/15 13:13:13
        0.0      0.0        
        """
        r = pe.get_sheet(file_name=os.path.join("tests", "fixtures", "date_field.xlsx"), library="openpyxl")
        assert isinstance(r[1,0], datetime.date) == True
        assert r[1,0].strftime("%d/%m/%y") == "25/12/14"
        assert isinstance(r[1,1], datetime.time) == True
        assert r[1,1].strftime("%H:%M:%S") == "11:11:11"
        value = r[4,0].isoformat()
        assert value == "1899-12-30T00:00:00"
        assert r[4,1].isoformat() == "1899-12-30T00:00:00"

    def test_writing_date_format(self):
        excel_filename = "testdateformat.xlsx"
        data = [[datetime.date(2014,12,25),
                datetime.time(11,11,11),
                 datetime.datetime(2014,12,25,11,11,11)]]
        pe.save_as(dest_file_name=excel_filename, array=data)
        r = pe.get_sheet(file_name=excel_filename, library="openpyxl")
        assert isinstance(r[0,0], datetime.date) == True
        assert r[0,0].strftime("%d/%m/%y") == "25/12/14"
        assert isinstance(r[0,1], datetime.time) == True
        assert r[0,1].strftime("%H:%M:%S") == "11:11:11"
        assert isinstance(r[0,2], datetime.date) == True
        assert r[0,2].strftime("%d/%m/%y") == "25/12/14"
        os.unlink(excel_filename)
