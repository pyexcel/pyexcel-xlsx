Change log
================================================================================

0.2.0 - 01.06.2016
--------------------------------------------------------------------------------

Added
++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

#. 'library=pyexcel-xlsx' was added to inform pyexcel to use it instead of other libraries, in the situation where there are more than one plugin for a file type, e.g. xlsm

Updated
++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

#. support the auto-import feature of pyexcel-io 0.2.0


0.1.0 - 17.01.2016
--------------------------------------------------------------------------------

Added
++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

#. Passing "streaming=True" to get_data, you will get the two dimensional array as a generator
#. Passing "data=your_generator" to save_data is acceptable too.

Updated
++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
#. compatibility with pyexcel-io 0.1.0
