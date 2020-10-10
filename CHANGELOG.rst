Change log
================================================================================

0.6.0 - 10.10.2020
--------------------------------------------------------------------------------

**Updated**

#. New style xlsx plugins, promoted by pyexcel-io v0.6.2.

0.5.8 - 28.12.2019
--------------------------------------------------------------------------------

**Updated**

#. `#34 <https://github.com/pyexcel/pyexcel-xlsx/issues/34>`_: pin
   openpyxl>=2.6.1

0.5.7 - 15.02.2019
--------------------------------------------------------------------------------

**Added**

#. `pyexcel-io#66 <https://github.com/pyexcel/pyexcel-io/issues/66>`_ pin
   openpyxl < 2.6.0

0.5.6 - 26.03.2018
--------------------------------------------------------------------------------

**Added**

#. `#24 <https://github.com/pyexcel/pyexcel-xlsx/issues/24>`_, remove deprecated
   warning from merged_cell_ranges and get_sheet_by_name

0.5.5 - 18.12.2017
--------------------------------------------------------------------------------

**Added**

#. `#22 <https://github.com/pyexcel/pyexcel-xlsx/issues/22>`_, to detect merged
   cell in xlsx - fast tracked patreon request.

0.5.4 - 2.11.2017
--------------------------------------------------------------------------------

**Updated**

#. Align the behavior of skip_hidden_row_and_column. Default it to True.

0.5.3 - 2.11.2017
--------------------------------------------------------------------------------

**Added**

#. `#20 <https://github.com/pyexcel/pyexcel-xlsx/issues/20>`_, skip hidden rows
   and columns under 'skip_hidden_row_and_column' flag.

0.5.2 - 23.10.2017
--------------------------------------------------------------------------------

**updated**

#. pyexcel `pyexcel#105 <https://github.com/pyexcel/pyexcel/issues/105>`_,
   remove gease from setup_requires, introduced by 0.5.1.
#. remove python2.6 test support
#. update its dependecy on pyexcel-io to 0.5.3

0.5.1 - 20.10.2017
--------------------------------------------------------------------------------

**added**

#. `pyexcel#103 <https://github.com/pyexcel/pyexcel/issues/103>`_, include
   LICENSE file in MANIFEST.in, meaning LICENSE file will appear in the released
   tar ball.

0.5.0 - 30.08.2017
--------------------------------------------------------------------------------

**Updated**

#. put dependency on pyexcel-io 0.5.0, which uses cStringIO instead of StringIO.
   Hence, there will be performance boost in handling files in memory.

**Removed**

#. `#18 <https://github.com/pyexcel/pyexcel-xlsx/issues/18>`_, is handled in
   pyexcel-io

0.4.2 - 25.08.2017
--------------------------------------------------------------------------------

**Updated**

#. `#18 <https://github.com/pyexcel/pyexcel-xlsx/issues/18>`_, handle unseekable
   stream given by http response

0.4.1 - 16.07.2017
--------------------------------------------------------------------------------

**Removed**

#. Removed useless code

0.4.0 - 19.06.2017
--------------------------------------------------------------------------------

**Updated**

#. `#14 <https://github.com/pyexcel/pyexcel-xlsx/issues/14>`_, close file handle
#. pyexcel-io plugin interface now updated to use `lml
   <https://github.com/chfw/lml>`_.

0.3.0 - 22.12.2016
--------------------------------------------------------------------------------

**Updated**

#. Code refactoring with pyexcel-io v 0.3.0
#. `#13 <https://github.com/pyexcel/pyexcel-xlsx/issues/13>`_, turn read_only
   flag on openpyxl.

0.2.3 - 05.11.2016
--------------------------------------------------------------------------------

**Updated**

#. `#12 <https://github.com/pyexcel/pyexcel-xlsx/issues/12>`_, remove
   UserWarning: Using a coordinate with ws.cell is deprecated. Use
   ws[coordinate]

0.2.2 - 31.08.2016
--------------------------------------------------------------------------------

**Added**

#. support pagination. two pairs: start_row, row_limit and start_column,
   column_limit help you deal with large files.

0.2.1 - 12.07.2016
--------------------------------------------------------------------------------

**Added**

#. `#8 <https://github.com/pyexcel/pyexcel-xlsx/issues/8>`__,
   `skip_hidden_sheets` is added. By default, hidden sheets are skipped when
   reading all sheets. Reading sheet by name or by index are not affected.

0.2.0 - 01.06.2016
--------------------------------------------------------------------------------

**Added**

#. 'library=pyexcel-xlsx' was added to inform pyexcel to use it instead of other
   libraries, in the situation where there are more than one plugin for a file
   type, e.g. xlsm

**Updated**

#. support the auto-import feature of pyexcel-io 0.2.0

0.1.0 - 17.01.2016
--------------------------------------------------------------------------------

**Added**

#. Passing "streaming=True" to get_data, you will get the two dimensional array
   as a generator
#. Passing "data=your_generator" to save_data is acceptable too.

**Updated**

#. compatibility with pyexcel-io 0.1.0
