A Python package to ease writing tables to Excel
================================================

|Status| |Docs|

Documentation here: https://xltable.readthedocs.org/en/latest

xltable is an API for writing tabular data and charts to Excel. It is not a replacement for other Excel writing
packages such as xlsxwriter, xlwt or pywin32. Instead it uses those packages as a back end to write the Excel files
(or to write to Excel directly in the case of pywin32) and provides a higer level abstraction that allows the
programmer to deal with tables of data rather than worry about writing individual cells.

The main feature that makes xltable more useful than just writing the Excel files directly is that it can
handle tables with formulas that relate to cells in the workbook without having to know in advance where
those tables will be placed on a worksheet. Only when all the tables have been added to the workbook and
the workbook is being written are formulas resolved to their final cell addresses.

Tables of data are constructed using pandas.DataFrame objects. These can contain formulas relating to columns or
cells in the same table or other tables in the same workbook.

As well as writing tables to Excel, xltable can also write charts using tables as source data.

Integrating xltable into Excel can be done using PyXLL, https://www.pyxll.com.
PyXLL embeds a Python interpreter within Excel and makes it possible to use Excel as a front end user interface
to Python code. For example, you could configure a custom ribbon control for users to run Python reports and have
the results written back to Excel.

Example::

    from xltable import *
    import pandas as pa

    # create a dataframe with three columns where the last is the sum of the first two
    dataframe = pa.DataFrame({
            "col_1": [1, 2, 3],
            "col_2": [4, 5, 6],
            "col_3": Cell("col_1") + Cell("col_2"),
    }, columns=["col_1", "col_2", "col_3"])

    # create the named xlwriter Table instance
    table = Table("table", dataframe)

    # create the Workbook and Worksheet objects and add table to the sheet
    sheet = Worksheet("Sheet1")
    sheet.add_table(table)

    workbook = Workbook("example.xlsx")
    workbook.add_sheet(sheet)

    # write the workbook to the file (requires xlsxwriter)
    workbook.to_xlsx()

.. |Status| image:: https://travis-ci.org/renshawbay/xltable.svg?branch=master
   :target: https://travis-ci.org/renshawbay/xltable

.. |Docs| image:: https://readthedocs.org/projects/xltable/badge/?version=latest
   :target: https://readthedocs.org/projects/xltable/?badge=latest
   :alt: Documentation Status
