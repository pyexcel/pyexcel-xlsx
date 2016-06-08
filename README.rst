================================================================================
pyexcel-xlsx - Let you focus on data, instead of xlsx format
================================================================================

.. image:: https://api.travis-ci.org/pyexcel/pyexcel-xlsx.png
    :target: http://travis-ci.org/pyexcel/pyexcel-xlsx

.. image:: https://codecov.io/github/pyexcel/pyexcel-xlsx/coverage.png
    :target: https://codecov.io/github/pyexcel/pyexcel-xlsx

**pyexcel-xlsx** is a tiny wrapper library to read, manipulate and write data in xlsx and xlsm fromat using openpyxl. You are likely to use it with `pyexcel <https://github.com/pyexcel/pyexcel>`__.

Known constraints
==================

Fonts, colors and charts are not supported.

Installation
================================================================================


Recently, pyexcel(0.2.2+) and its plugins(0.2.0+) started using newer version of setuptools. Please upgrade your setup tools before install latest pyexcel components:

.. code-block:: bash

    $ pip install --upgrade setuptools

You can install it via pip:

.. code-block:: bash

    $ pip install pyexcel-xlsx


or clone it and install it:

.. code-block:: bash

    $ git clone http://github.com/pyexcel/pyexcel-xlsx.git
    $ cd pyexcel-xlsx
    $ python setup.py install

Usage
================================================================================

As a standalone library
--------------------------------------------------------------------------------

Write to an xlsx file
********************************************************************************

.. testcode::
   :hide:

    >>> import sys
    >>> if sys.version_info[0] < 3:
    ...     from StringIO import StringIO
    ... else:
    ...     from io import BytesIO as StringIO
    >>> PY2 = sys.version_info[0] == 2
    >>> if PY2 and sys.version_info[1] < 7:
    ...      from ordereddict import OrderedDict
    ... else:
    ...     from collections import OrderedDict


Here's the sample code to write a dictionary to an xlsx file:

.. code-block:: python

    >>> from pyexcel_xlsx import save_data
    >>> data = OrderedDict() # from collections import OrderedDict
    >>> data.update({"Sheet 1": [[1, 2, 3], [4, 5, 6]]})
    >>> data.update({"Sheet 2": [["row 1", "row 2", "row 3"]]})
    >>> save_data("your_file.xlsx", data)

Read from an xlsx file
********************************************************************************

Here's the sample code:

.. code-block:: python

    >>> from pyexcel_xlsx import get_data
    >>> data = get_data("your_file.xlsx")
    >>> import json
    >>> print(json.dumps(data))
    {"Sheet 1": [[1, 2, 3], [4, 5, 6]], "Sheet 2": [["row 1", "row 2", "row 3"]]}

Write an xlsx to memory
********************************************************************************

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
********************************************************************************

Continue from previous example:

.. code-block:: python

    >>> # This is just an illustration
    >>> # In reality, you might deal with xlsx file upload
    >>> # where you will read from requests.FILES['YOUR_XLSX_FILE']
    >>> data = get_data(io)
    >>> print(json.dumps(data))
    {"Sheet 1": [[1, 2, 3], [4, 5, 6]], "Sheet 2": [[7, 8, 9], [10, 11, 12]]}


As a pyexcel plugin
--------------------------------------------------------------------------------

No longer, explicit import is needed since pyexcel version 0.2.2. Instead,
this library is auto-loaded. So if you want to read data in xlsx format,
installing it is enough.

Any version under pyexcel 0.2.2, you have to keep doing the following:

Import it in your file to enable this plugin:

.. code-block:: python

    from pyexcel.ext import xlsx

Please note only pyexcel version 0.0.4+ support this.

Reading from an xlsx file
********************************************************************************

Here is the sample code:

.. code-block:: python

    >>> import pyexcel as pe
    >>> # from pyexcel.ext import xlsx
    >>> sheet = pe.get_book(file_name="your_file.xlsx")
    >>> sheet
    Sheet 1:
    +---+---+---+
    | 1 | 2 | 3 |
    +---+---+---+
    | 4 | 5 | 6 |
    +---+---+---+
    Sheet 2:
    +-------+-------+-------+
    | row 1 | row 2 | row 3 |
    +-------+-------+-------+

Writing to an xlsx file
********************************************************************************

Here is the sample code:

.. code-block:: python

    >>> sheet.save_as("another_file.xlsx")

Reading from a IO instance
================================================================================

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
    Sheet 1:
    +---+---+---+
    | 1 | 2 | 3 |
    +---+---+---+
    | 4 | 5 | 6 |
    +---+---+---+
    Sheet 2:
    +-------+-------+-------+
    | row 1 | row 2 | row 3 |
    +-------+-------+-------+


Writing to a StringIO instance
================================================================================

You need to pass a StringIO instance to Writer:

.. code-block:: python

    >>> data = [
    ...     [1, 2, 3],
    ...     [4, 5, 6]
    ... ]
    >>> io = StringIO()
    >>> sheet = pe.Sheet(data)
    >>> io = sheet.save_to_memory("xlsx", io)
    >>> # then do something with io
    >>> # In reality, you might give it to your http response
    >>> # object for downloading

License
================================================================================

New BSD License

Developer guide
==================

Development steps for code changes

#. git clone https://github.com/pyexcel/pyexcel-xlsx.git
#. cd pyexcel-xlsx
#. pip install -r rnd_requirements.txt # if such a file exists
#. pip install -r requirements.txt
#. pip install -r tests/requirements.txt


In order to update test envrionment, and documentation, additional setps are
required:

#. pip install moban
#. git clone https://github.com/pyexcel/pyexcel-commons.git
#. make your changes in `.moban.d` directory, then issue command `moban`

What is rnd_requirements.txt
-------------------------------

Usually, it is created when a depdent library is not released. Once the dependecy is installed(will be released), the future version of the dependency in the requirements.txt will be valid.

What is pyexcel-commons
---------------------------------

Many information that are shared across pyexcel projects, such as: this developer guide, license info, etc. are stored in `pyexcel-commons` project.

What is .moban.d
---------------------------------

`.moban.d` stores the specific meta data for the library.

How to test your contribution
------------------------------

Although `nose` and `doctest` are both used in code testing, it is adviable that unit tests are put in tests. `doctest` is incorporated only to make sure the code examples in documentation remain valid across different development releases.

On Linux/Unix systems, please launch your tests like this::

    $ make test

On Windows systems, please issue this command::

    > test.bat


.. testcode::
   :hide:

   >>> import os
   >>> os.unlink("your_file.xlsx")
   >>> os.unlink("another_file.xlsx")
