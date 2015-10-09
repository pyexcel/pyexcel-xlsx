==============================================================
pyexcel-xlsx - Let you focus on data, instead of xlsx format
==============================================================

.. image:: https://travis-ci.org/chfw/pyexcel-xlsx.svg
    :target: https://travis-ci.org/chfw/pyexcel-xlsx

.. image:: https://coveralls.io/repos/chfw/pyexcel-xlsx/badge.png?branch=master
    :target: https://coveralls.io/r/chfw/pyexcel-xlsx?branch=master

**pyexcel-xlsx** is a tiny wrapper library to read, manipulate and write data in xlsx and xlsm fromat using openpyxl. You are likely to use it with `pyexcel <https://github.com/chfw/pyexcel>`__. 

Known constraints
==================

Fonts, colors and charts are not supported. 

Installation
============

You can install it via pip:

.. code-block:: bash

    $ pip install pyexcel-xlsx


or clone it and install it:

.. code-block:: bash

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
    >>> from pyexcel_io import OrderedDict

Here's the sample code to write a dictionary to an xlsx file:

.. code-block:: python

    >>> from pyexcel_xlsx import save_data
    >>> data = OrderedDict() # from collections import OrderedDict
    >>> data.update({"Sheet 1": [[1, 2, 3], [4, 5, 6]]})
    >>> data.update({"Sheet 2": [["row 1", "row 2", "row 3"]]})
    >>> save_data("your_file.xlsx", data)

Read from an xlsx file
**********************

Here's the sample code:

.. code-block:: python

    >>> from pyexcel_xlsx import get_data
    >>> data = get_data("your_file.xlsx")
    >>> import json
    >>> print(json.dumps(data))
    {"Sheet 1": [[1, 2, 3], [4, 5, 6]], "Sheet 2": [["row 1", "row 2", "row 3"]]}

Write an xlsx to memory
*************************

Here's the sample code to write a dictionary to an xlsx file:

.. code-block:: python

    >>> from pyexcel_xlsx import save_data
    >>> data = OrderedDict()
    >>> data.update({"Sheet 1": [[1, 2, 3], [4, 5, 6]]})
    >>> data.update({"Sheet 2": [[7, 8, 9], [10, 11, 12]]})
    >>> io = StringIO()
    >>> save_data(io, data)
    >>> # do something with the io
    >>> # In reality, you might give it to your http response
    >>> # object for downloading

    
Read from an xlsx from memory
*****************************

Continue from previous example:

.. code-block:: python

    >>> # This is just an illustration
    >>> # In reality, you might deal with xlsx file upload
    >>> # where you will read from requests.FILES['YOUR_XLSX_FILE']
    >>> data = get_data(io)
    >>> print(json.dumps(data))
    {"Sheet 1": [[1, 2, 3], [4, 5, 6]], "Sheet 2": [[7, 8, 9], [10, 11, 12]]}


As a pyexcel plugin
--------------------

Import it in your file to enable this plugin:

.. code-block:: python

    from pyexcel.ext import xlsx

Please note only pyexcel version 0.0.4+ support this.

Reading from an xlsx file
**************************

Here is the sample code:

.. code-block:: python

    >>> import pyexcel as pe
    >>> from pyexcel.ext import xlsx
    >>> sheet = pe.get_book(file_name="your_file.xlsx")
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
************************

Here is the sample code:

.. code-block:: python

    >>> sheet.save_as("another_file.xlsx")

Reading from a IO instance
================================

You got to wrap the binary content with stream to get xlsx working:

.. code-block:: python

    >>> # This is just an illustration
    >>> # In reality, you might deal with xlsx file upload
    >>> # where you will read from requests.FILES['YOUR_XLSX_FILE']
    >>> xlsxfile = "another_file.xlsx"
    >>> with open(xlsxfile, "rb") as f:
    ...     content = f.read()
    ...     r = pe.get_book(file_type="xlsx", file_content=content)
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

You need to pass a StringIO instance to Writer:

.. code-block:: python

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

License
==========

New BSD License

Dependencies
============

1. openpyxl
2. pyexcel-io >= 0.0.8

.. testcode::
   :hide:

   >>> import os
   >>> os.unlink("your_file.xlsx")
   >>> os.unlink("another_file.xlsx")
