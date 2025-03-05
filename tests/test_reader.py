import os
from datetime import time, datetime

from pyexcel_xlsx import get_data
from pyexcel_io._compact import OrderedDict

from nose.tools import eq_


def test_reading():
    data = get_data(
        os.path.join("tests", "fixtures", "date_field.xlsx"),
        library="pyexcel-xlsx",
        skip_hidden_row_and_column=False,
    )
    expected = OrderedDict()
    expected.update(
        {
            "Sheet1": [
                ["Date", "Time"],
                [
                    datetime(year=2014, month=12, day=25),
                    time(hour=11, minute=11, second=11),
                ],
                [
                    datetime(2014, 12, 26, 0, 0),
                    time(hour=12, minute=12, second=12),
                ],
                [
                    datetime(2015, 1, 1, 0, 0),
                    time(hour=13, minute=13, second=13),
                ],
                [time(0, 0), time(0, 0)],
            ]
        }
    )
    expected.update({"Sheet2": []})
    expected.update({"Sheet3": []})
    eq_(data, expected)
