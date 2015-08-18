"""
xltable
=======

A package for writing excel worksheets with formulas and styles.

Introduction
------------

:py:mod:`xltable` is an API for writing tabular data and charts to Excel.
It is not a replacement for other Excel writing packages such as :py:mod:`xlsxwriter`,
:py:mod:`xlwt` or :py:mod:`pywin32`. Instead it uses those packages as a back end
to write the Excel files (or to write to Excel directly in the case of :py:mod:`pywin32`)
and provides a higer level abstraction that allows the programmer to deal with
tables of data rather than worry about writing individual cells.

The main feature that makes :py:mod:`xltable` more useful than just writing the
Excel files directly is that it can handle tables with formulas that relate to
cells in the workbook *without* having to know in advance where those tables
will be placed on a worksheet. Only when all the tables have been added to the
workbook and the workbook is being written are formulas resolved to their
final cell addresses.

Tables of data are constructed using :py:class:`pandas.DataFrame` objects. These
can contain formulas relating to columns or cells in the same table or other
tables in the same workbook.

As well as writing tables to Excel, :py:mod:`xltable` can also write charts using
tables as source data.

Example
-------

Write a dataframe with a formula to Excel::

    from xltable import *
    import pandas as pa

    # create a dataframe with three columns where the last is the sum of the first two
    dataframe = pa.DataFrame({
            "col_1": [1, 2, 3],
            "col_2": [4, 5, 6],
            "col_3": Cell("col_1") + Cell("col_2"),
    }, columns=["col_1", "col_2", "col_3"])

    # create the named xltable Table instance
    table = Table("table", dataframe)

    # create the Workbook and Worksheet objects and add table to the sheet
    sheet = Worksheet("Sheet1")
    sheet.add_table(table)

    workbook = Workbook("example.xlsx")
    workbook.add_sheet(sheet)

    # write the workbook to the file (requires xlsxwriter)
    workbook.to_xlsx()

Classes
-------

.. autoclass:: Workbook

    .. automethod:: add_sheet

    .. automethod:: to_xlsx

    .. automethod:: to_excel


.. autoclass:: Worksheet

    .. autoattribute:: name

    .. automethod:: add_table

    .. automethod:: add_chart

    .. automethod:: get_table_pos

    .. automethod:: get_table

    .. autoattribute:: next_row

.. autoclass:: Table

    .. automethod:: get_data

.. autoclass:: Chart

    .. automethod:: add_series

.. autoclass:: ArrayFormula

.. autoclass:: Expression

.. autoclass:: Cell

.. autoclass:: Column

.. autoclass:: Index

.. autoclass:: Range

.. autoclass:: Formula

.. autoclass:: TableStyle

.. autoclass:: CellStyle

.. autoclass:: Value

"""
from .expression import Column, Index, Cell, Range, Formula, ConstExpr, Expression
from .style import CellStyle, TableStyle
from .table import Table, Value, ArrayFormula
from .chart import Chart
from .worksheet import Worksheet
from .workbook import Workbook

__all__ = [
    "Workbook",
    "Worksheet",
    "Table",
    "Chart",
    "Value",
    "Formula",
    "ArrayFormula",
    "CellStyle",
    "TableStyle",
    "Column",
    "Index",
    "Expression",
    "Cell",
    "Range",
    "Formula",
    "ConstExpr",
]
