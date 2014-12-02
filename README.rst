===========
pyexcel-xl
===========

.. image:: https://travis-ci.org/chfw/pyexcel-xlsx.svg?branch=v0.0.1-rc1
    :target: https://travis-ci.org/chfw/pyexcel-xlsx/builds/41641140

.. image:: https://coveralls.io/repos/chfw/pyexcel-xlsx/badge.png?branch=v0.0.1-rc1 
    :target: https://coveralls.io/r/chfw/pyexcel-xlsx?branch=v0.0.1-rc1 

.. image:: https://pypip.in/d/pyexcel-xlsx/badge.png
    :target: https://pypi.python.org/pypi/pyexcel-xlsx

.. image:: https://pypip.in/py_versions/pyexcel-xlsx/badge.png
    :target: https://pypi.python.org/pypi/pyexcel-xlsx

.. image:: https://pypip.in/implementation/pyexcel-xlsx/badge.png
    :target: https://pypi.python.org/pypi/pyexcel-xlsx

**pyexcel-xlsx** is a tiny wrapper library to read, manipulate and write data in xls, xlsx and xlsm fromat. You are likely to use it with `pyexcel <https://github.com/chfw/pyexcel>`__. 

Installation
============

You can install it via pip::

    $ pip install pyexcel-xlsx


or clone it and install it::

    $ git clone http://github.com/chfw/pyexcel-xlsx.git
    $ cd pyexcel-xlsx
    $ python setup.py install

Usage
=====


As a standalone library
------------------------

Write to an xlsx file
*********************

.. testcode::
   :hide:

    >>> import sys
    >>> if sys.version_info[0] < 3:
    ...     from StringIO import StringIO
    ... else:
    ...     from io import BytesIO as StringIO
    >>> from pyexcel_xlsx.xlbook import OrderedDict

Here's the sample code to write a dictionary to an xlsx file::

.. code:: python

    >>> from pyexcel_xlsx import XLSXWriter
    >>> data = OrderedDict() # from collections import OrderedDict
    >>> data.update({"Sheet 1": [[1, 2, 3], [4, 5, 6]]})
    >>> data.update({"Sheet 2": [["row 1", "row 2", "row 3"]]})
    >>> writer = XLSXWriter("your_file.xlsx")
    >>> writer.write(data)
    >>> writer.close()

Read from an xlsx file
**********************

Here's the sample code::

.. code:: python

    >>> from pyexcel_xlsx import XLSXBook

    >>> book = XLSXBook("your_file.xlsx")
    >>> # book.sheets() returns a dictionary of all sheet content
    >>> #   the keys represents sheet names
    >>> #   the values are two dimensional array
    >>> import json
    >>> print(json.dumps(book.sheets()))
    {"Sheet 1": [[1, 2, 3], [4, 5, 6]], "Sheet 2": [["row 1", "row 2", "row 3"]]}

Write an xlsx to memory
**********************

Here's the sample code to write a dictionary to an xlsx file::

.. code:: python

    >>> from pyexcel_xlsx import XLSXWriter
    >>> data = OrderedDict()
    >>> data.update({"Sheet 1": [[1, 2, 3], [4, 5, 6]]})
    >>> data.update({"Sheet 2": [[7, 8, 9], [10, 11, 12]]})
    >>> io = StringIO()
    >>> writer = XLSXWriter(io)
    >>> writer.write(data)
    >>> writer.close()
    >>> # do something witht the io
    >>> # In reality, you might give it to your http response
    >>> # object for downloading

    
Read from an xlsx from memory
*****************************

Continue from previous example::

.. code:: python

    >>> # This is just an illustration
    >>> # In reality, you might deal with xlsx file upload
    >>> # where you will read from requests.FILES['YOUR_XLSX_FILE']
    >>> book = XLSXBook(None, io.getvalue())
    >>> print(json.dumps(book.sheets()))
    {"Sheet 1": [[1, 2, 3], [4, 5, 6]], "Sheet 2": [[7, 8, 9], [10, 11, 12]]}


As a pyexcel plugin
--------------------

Import it in your file to enable this plugin::

    from pyexcel.ext import xlsx

Please note only pyexcel version 0.0.4+ support this.

Reading from an xlsx file
************************

Here is the sample code::

.. code:: python

    >>> import pyexcel as pe
    >>> from pyexcel.ext import xlsx
    
    # "example.xlsx"
    >>> sheet = pe.load_book("your_file.xlsx")
    >>> sheet
    Sheet Name: Sheet 1
    +---+---+---+
    | 1 | 2 | 3 |
    +---+---+---+
    | 4 | 5 | 6 |
    +---+---+---+
    Sheet Name: Sheet 2
    +-------+-------+-------+
    | row 1 | row 2 | row 3 |
    +-------+-------+-------+

Writing to an xlsx file
**********************

Here is the sample code::

.. code:: python

    >>> sheet.save_as("another_file.xlsx")

Reading from a IO instance
================================

You got to wrap the binary content with stream to get xlsx working::

.. code:: python

    >>> # This is just an illustration
    >>> # In reality, you might deal with xlsx file upload
    >>> # where you will read from requests.FILES['YOUR_XLSX_FILE']
    >>> xlsxfile = "another_file.xlsx"
    >>> with open(xlsxfile, "rb") as f:
    ...     content = f.read()
    ...     r = pe.load_book_from_memory("xlsx", content)
    ...     print(r)
    ...
    Sheet Name: Sheet 1
    +---+---+---+
    | 1 | 2 | 3 |
    +---+---+---+
    | 4 | 5 | 6 |
    +---+---+---+
    Sheet Name: Sheet 2
    +-------+-------+-------+
    | row 1 | row 2 | row 3 |
    +-------+-------+-------+


Writing to a StringIO instance
================================

You need to pass a StringIO instance to Writer::

.. code:: python

    >>> data = [
    ...     [1, 2, 3],
    ...     [4, 5, 6]
    ... ]
    >>> io = StringIO()
    >>> sheet = pe.Sheet(data)
    >>> sheet.save_to_memory("xlsx", io)
    >>> # then do something with io
    >>> # In reality, you might give it to your http response
    >>> # object for downloading


Dependencies
============

1. openpyxl

.. testcode::
   :hide:

   >>> import os
   >>> os.unlink("your_file.xlsx")
   >>> os.unlink("another_file.xlsx")

