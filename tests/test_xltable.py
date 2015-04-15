import unittest
import pandas as pa
from xltable import *


class XLTableTest(unittest.TestCase):

    def test_simple_table(self):
        """test writing a couple of dataframes to a worksheet"""
        worksheet = Worksheet()

        df = pa.DataFrame({
            "A": [1, 2, 3],
            "B": [4, 5, 6],
            "C": [7, 8, 9],
        }, columns=["A", "B", "C"])

        table_1 = Table("table_1", dataframe=df)
        worksheet.add_table(table_1)

        df = pa.DataFrame({
            "X": [1, 2, 3],
            "Y": [4, 5, 6],
        }, columns=["X", "Y"])

        table_2 = Table("table_2", dataframe=df)
        worksheet.add_table(table_2)

        row_iter = iter(worksheet.iterrows())

        # first table
        row = next(row_iter)
        self.assertEqual(row, ["A", "B", "C"])
        row = next(row_iter)
        self.assertEqual(row, [1, 4, 7])
        row = next(row_iter)
        self.assertEqual(row, [2, 5, 8])
        row = next(row_iter)
        self.assertEqual(row, [3, 6, 9])

        # space between tables
        row = next(row_iter)
        self.assertEqual(row, [None, None, None])

        # second table
        row = next(row_iter)
        self.assertEqual(row, ["X", "Y", None])
        row = next(row_iter)
        self.assertEqual(row, [1, 4, None])
        row = next(row_iter)
        self.assertEqual(row, [2, 5, None])
        row = next(row_iter)
        self.assertEqual(row, [3, 6, None])

        # end of sheet
        self.assertRaises(StopIteration, next, row_iter)

    def test_formula(self):
        """test formulas get resolved correctly"""
        workbook = Workbook()

        df_1 = pa.DataFrame({
            "col_1": [1, 2, 3],
            "col_2": [4, 5, 6],
            "col_3": Formula("SUM", Cell("col_1"), Cell("col_2")),
        }, columns=["col_1", "col_2", "col_3"])

        table_1 = Table("table_1", df_1)
        sheet_1 = Worksheet("Sheet1")
        sheet_1.add_table(table_1)
        workbook.add_sheet(sheet_1)

        df_2 = pa.DataFrame({
            "SP": Formula("SUMPRODUCT", Column("col_1", table="table_1"), Column("col_2", table="table_1"))
        }, index=[1], columns=["SP"])

        table_2 = Table("table_2", df_2)
        sheet_2 = Worksheet("Sheet2")
        sheet_2.add_table(table_2)
        workbook.add_sheet(sheet_2)

        # table 1's third column should be a formula relating to the first two columns
        row, col = sheet_1.get_table_pos("table_1")
        data = table_1.get_data(workbook, row, col)

        self.assertEqual(data.shape, (4, 3))
        self.assertEqual(data[0].tolist(), ["col_1", "col_2", "col_3"])
        self.assertEqual(data[1].tolist(), [1, 4, "=SUM('Sheet1'!A2,'Sheet1'!B2)"])
        self.assertEqual(data[2].tolist(), [2, 5, "=SUM('Sheet1'!A3,'Sheet1'!B3)"])
        self.assertEqual(data[3].tolist(), [3, 6, "=SUM('Sheet1'!A4,'Sheet1'!B4)"])

        # table 2 should refer to table 1, on Sheet1
        row, col = sheet_2.get_table_pos("table_2")
        data = table_2.get_data(workbook, row, col)
        self.assertEqual(data.shape, (2, 1))

        self.assertEqual(data[0].tolist(), ["SP"])
        self.assertEqual(data[1].tolist(), ["=SUMPRODUCT('Sheet1'!$A$2:$A$4,'Sheet1'!$B$2:$B$4)"])
