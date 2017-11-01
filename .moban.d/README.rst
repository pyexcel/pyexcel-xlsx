{%extends 'README.rst.jj2' %}

{%block description%}
**pyexcel-xlsx** is a tiny wrapper library to read, manipulate and write data in xlsx and xlsm fromat using openpyxl. You are likely to use it with `pyexcel <https://github.com/pyexcel/pyexcel>`__.

Please note:

1. `auto_detect_int` flag will not take effect because openpyxl detect integer in python 3 by default.
2. `skip_hidden_row_and_column` will slow down the read operation.


{%endblock%}
