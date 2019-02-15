import os

from pyexcel_xlsx import get_data

from nose.tools import eq_


def test_hidden_row():
    data = get_data(
        os.path.join("tests", "fixtures", "hidden.xlsx"),
        skip_hidden_row_and_column=True,
        library="pyexcel-xlsx",
    )
    expected = [[1, 2], [7, 9]]
    eq_(data["Sheet1"], expected)


def test_complex_hidden_sheets():
    data = get_data(
        os.path.join("tests", "fixtures", "complex_hidden_sheets.xlsx"),
        skip_hidden_row_and_column=True,
        library="pyexcel-xlsx",
    )
    expected = [[1, 3, 5, 7, 9], [31, 33, 35, 37, 39], [61, 63, 65, 67]]
    eq_(data["Sheet1"], expected)
