Change log
================================================================================

0.4.0 - 19.06.2017
--------------------------------------------------------------------------------

Updated
********************************************************************************

#. `#14 <https://github.com/pyexcel/pyexcel-xlsx/issues/14>`_, close file
   handle
#. pyexcel-io plugin interface now updated to use
   `lml <https://github.com/chfw/lml>`_.

0.3.0 - 22.12.2016
--------------------------------------------------------------------------------

Updated
********************************************************************************

#. Code refactoring with pyexcel-io v 0.3.0
#. `#13 <https://github.com/pyexcel/pyexcel-xlsx/issues/13>`_, turn read_only
   flag on openpyxl.

0.2.3 - 05.11.2016
--------------------------------------------------------------------------------

Updated
********************************************************************************

#. `#12 <https://github.com/pyexcel/pyexcel-xlsx/issues/12>`_, remove
   UserWarning: Using a coordinate with ws.cell is deprecated.
   Use ws[coordinate]


0.2.2 - 31.08.2016
--------------------------------------------------------------------------------

Added
********************************************************************************

#. support pagination. two pairs: start_row, row_limit and start_column, column_limit
   help you deal with large files.


0.2.1 - 12.07.2016
--------------------------------------------------------------------------------

Added
********************************************************************************

#. `#8 <https://github.com/pyexcel/pyexcel-xlsx/issues/8>`__, `skip_hidden_sheets` is added. By default, hidden sheets are skipped when reading all sheets. Reading sheet by name or by index are not affected.


0.2.0 - 01.06.2016
--------------------------------------------------------------------------------

Added
********************************************************************************

#. 'library=pyexcel-xlsx' was added to inform pyexcel to use it instead of other libraries, in the situation where there are more than one plugin for a file type, e.g. xlsm

Updated
********************************************************************************

#. support the auto-import feature of pyexcel-io 0.2.0


0.1.0 - 17.01.2016
--------------------------------------------------------------------------------

Added
********************************************************************************

#. Passing "streaming=True" to get_data, you will get the two dimensional array as a generator
#. Passing "data=your_generator" to save_data is acceptable too.

Updated
********************************************************************************
#. compatibility with pyexcel-io 0.1.0
