"""
Collection of worksheet instances
"""
import logging

_log = logging.getLogger(__name__)


class Workbook(object):
    """
    A workbook is an ordered collection of worksheets.
    
    Once all worksheets have been added the workbook can be written out or
    the worksheets can be iterated over, and any expressions present in the
    tables of the worksheets will be resolved to absolute worksheet/cell references.

    :param str filename: Filename the workbook will be written to.
    :param list worksheets: List of :py:class:`xltable.Worksheet` instances.
    """
    def __init__(self, filename=None, worksheets=[]):
        self.filename = filename
        self.worksheets = list(worksheets)
        self.calc_mode = "auto"
        self.workbook_obj = None

        # The active table and worksheet objects are set during export, and
        # are used to resolve expressions where the table and/or sheet isn't
        # set explicitly (in which case the current table is used implicitly).
        self.active_table = None
        self.active_worksheet = None

    def add_sheet(self, worksheet):
        """
        Adds a worksheet to the workbook.
        """
        self.worksheets.append(worksheet)

    # alias for add_sheet
    append = add_sheet

    def set_calc_mode(self, mode):
        """
        Set the calculation mode for the Excel workbook
        """
        self.calc_mode = mode

    def itersheets(self):
        """
        Iterates over the worksheets in the book, and sets the active
        worksheet as the current one before yielding.
        """
        for ws in self.worksheets:
            # Expression with no explicit table specified will use None
            # when calling get_table, which should return the current worksheet/table
            prev_ws = self.active_worksheet
            self.active_worksheet = ws
            try:
                yield ws
            finally:
                self.active_worksheet = prev_ws

    def to_xlsx(self, **kwargs):
        """
        Write workbook to a .xlsx file using xlsxwriter.
        Return a xlsxwriter.workbook.Workbook.

        :param kwargs: Extra arguments passed to the xlsxwriter.Workbook
        constructor.
        """
        from xlsxwriter.workbook import Workbook as _Workbook
        self.workbook_obj = _Workbook(**kwargs)
        self.workbook_obj.set_calc_mode(self.calc_mode)

        for worksheet in self.itersheets():
            worksheet.to_xlsx(workbook=self)

        self.workbook_obj.filename = self.filename
        if self.filename:
            self.workbook_obj.close()
        return self.workbook_obj

    def to_excel(self, xl_app=None, resize_columns=True):
        from win32com.client import Dispatch, gencache

        if xl_app is None:
            xl_app = Dispatch("Excel.Application")
        xl_app = gencache.EnsureDispatch(xl_app)

        # Add a new workbook with the correct number of sheets.
        # We aren't allowed to create an empty one.
        assert self.worksheets, "Can't export workbook with no worksheets"
        sheets_in_new_workbook = xl_app.SheetsInNewWorkbook
        try:
            xl_app.SheetsInNewWorkbook = float(len(self.worksheets))
            self.workbook_obj = xl_app.Workbooks.Add()
        finally:
            xl_app.SheetsInNewWorkbook = sheets_in_new_workbook

        # Rename the worksheets, ensuring that there can never be two sheets with the same
        # name due to the sheets default names conflicting with the new names.
        sheet_names = {s.name for s in self.worksheets}
        assert len(sheet_names) == len(self.worksheets), "Worksheets must have unique names"
        for worksheet in self.workbook_obj.Sheets:
            i = 1
            original_name = worksheet.Name
            while worksheet.Name in sheet_names:
                worksheet.Name = "%s_%d" % (original_name, i)
                i += 1

        for worksheet, sheet in zip(self.workbook_obj.Sheets, self.worksheets):
            worksheet.Name = sheet.name

        # Export each sheet (have to use itersheets for this as it sets the
        # current active sheet before yielding each one).
        for worksheet, sheet in zip(self.workbook_obj.Sheets, self.itersheets()):
            worksheet.Select()
            sheet.to_excel(workbook=self,
                           worksheet=worksheet,
                           xl_app=xl_app,
                           rename=False,
                           resize_columns=resize_columns)

        return self.workbook_obj

    def get_last_sheet(self):
        return self.workbook_obj.Sheets[self.workbook_obj.Sheets.Count]

    def add_xlsx_worksheet(self, worksheet, name):
        if worksheet not in self.worksheets:
            self.append(worksheet)
        return self.workbook_obj.add_worksheet(name)

    def add_excel_worksheet(self, after=None):
        if after is None:
            after = self.get_last_sheet()
        return self.workbook_obj.Sheets.Add(After=after)

    def add_format(self, *args, **kwargs):
        return self.workbook_obj.add_format(*args, **kwargs)

    def get_table(self, name):
        """
        Return a table, worksheet pair for the named table
        """
        if name is None:
            assert self.active_table, "Can't get table without name unless an active table is set"
            name = self.active_table.name

            if self.active_worksheet:
                table = self.active_worksheet.get_table(name)
                assert table is self.active_table, "Active table is not from the active sheet"
                return table, self.active_worksheet

            for ws in self.worksheets:
                try:
                    table = ws.get_table(name)
                    if table is self.active_table:
                        return table, ws
                except KeyError:
                    pass

            raise RuntimeError("Active table not found in any sheet")

        # if the tablename explicitly uses the sheetname find the right sheet
        if "!" in name:
            ws_name, table_name = map(lambda x: x.strip("'"), name.split("!", 1))
            for ws in self.worksheets:
                if ws.name == ws_name:
                    table = ws.get_table(table_name)
                    return table, ws
            raise KeyError(name)

        # otherwise look in the current table
        if self.active_worksheet:
            table = self.active_worksheet.get_table(name)
            return table, self.active_worksheet

        # or fallback to the first matching name in any table
        for ws in self.worksheets:
            try:
                table = ws.get_table(name)
                return table, ws
            except KeyError:
                pass

        raise KeyError(name)
