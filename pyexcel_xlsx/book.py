"""
pyexcel_xlsx.xlsxr
~~~~~~~~~~~~~~~~~~~

Read xlsx file format using openpyxl

:copyright: (c) 2015-2025 by Onni Software Ltd & its contributors
:license: New BSD License
"""

import openpyxl


class Sheet(object):
    def __init__(self, xlsx_sheet):
        self.xlsx_sheet = xlsx_sheet

    def __getitem__(self, cell):
        return self.xlsx_sheet[cell]

    def __setitem__(self, cell, value):
        self.xlsx_sheet[cell] = value

    def append(self, data):
        self.xlsx_sheet.append(data)


class Book(object):
    def __init__(self, file_name):
        self.xlsx_book = openpyxl.load_workbook(
            filename=file_name,
        )

    def save(self, file_name):
        self.xlsx_book.save(file_name)

    def close(self):
        self.xlsx_book.close()

    def __getitem__(self, key):
        if key not in self.xlsx_book.sheetnames:
            self.xlsx_book.create_sheet(key)
        return Sheet(self.xlsx_book[key])

    def __delitem__(self, other):
        self.xlsx_book.remove(self.xlsx_book[other])
