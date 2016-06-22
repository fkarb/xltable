"""
A worksheet is a collection of tables placed at specific locations.

Once all tables have been placed the worksheet can be written out or
the rows can be iterated over, and any expressions present in the
tables will be resolved to absolute cell references.
"""
from .style import CellStyle
from .table import ArrayFormula, Value
from .expression import Expression
import re
import datetime as dt
import pandas as pa
import numpy as np
from copy import copy


class Worksheet(object):
    """
    A worksheet is a collection of tables placed at specific locations.
    
    Once all tables have been placed the worksheet can be written out or
    the rows can be iterated over, and any expressions present in the
    tables will be resolved to absolute cell references.

    :param str name: Worksheet name.
    """
    _xlsx_unsupported_types = tuple()

    def __init__(self, name="Sheet1"):
        self.__name = name
        self.__tables = {}
        self.__values = {}
        self.__charts = []
        self.__next_row = 0
        self.__groups = []

    @property
    def name(self):
        """Worksheet name"""
        return self.__name

    def add_table(self, table, row=None, col=0, row_spaces=1):
        """
        Adds a table to the worksheet at (row, col).
        Return the (row, col) where the table has been put.

        :param xltable.Table table: Table to add to the worksheet.
        :param int row: Row to start the table at (defaults to the next free row).
        :param int col: Column to start the table at.
        :param int row_spaces: Number of rows to leave between this table and the next.
        """
        name = table.name
        assert name is not None, "Tables must have a name"
        assert name not in self.__tables, "Table %s already exists in this worksheet" % name
        if row is None:
            row = self.__next_row
        self.__next_row = max(row + table.height + row_spaces, self.__next_row)
        self.__tables[name] = (table, (row, col))
        return row, col

    def add_value(self, value, row, col):
        """
        Adds a single value (cell) to a worksheet at (row, col).
        Return the (row, col) where the value has been put.

        :param value: Value to write to the sheet.
        :param row: Row where the value should be written.
        :param col: Column where the value should be written.
        """
        self.__values[(row, col)] = value

    def add_chart(self, chart, row, col):
        """
        Adds a chart to the worksheet at (row, col).

        :param xltable.Chart Chart: chart to add to the workbook.
        :param int row: Row to add the chart at.
        """
        self.__charts.append((chart, (row, col)))

    def add_row_group(self, tables, collapsed=True):
        """
        Adds a group over all the given tables (will include any rows between the first row over all
        tables, and the last row over all tables)
        Initially collapsed if collapsed is True (True by default)
        """
        self.__groups.append((tables, collapsed))

    @property
    def next_row(self):
        """Row the next table will start at unless another row is specified."""
        return self.__next_row

    @next_row.setter
    def next_row(self, value):
        self.__next_row = value

    def get_table_pos(self, tablename):
        """
        :param str tablename: Name of table to get position of.
        :return: Upper left (row, col) coordinate of the named table.
        """
        _table, (row, col) = self.__tables[tablename]
        return (row, col)

    def get_table(self, tablename):
        """
        :param str tablename: Name of table to find.
        :return: A :py:class:`xltable.Table` instance from the table name.
        """
        table, (_row, _col) = self.__tables[tablename]
        return table

    def iterrows(self, workbook=None):
        """
        Yield rows as lists of data.

        The data is exactly as it is in the source pandas DataFrames and
        any formulas are not resolved.
        """
        resolved_tables = []
        max_height = 0
        max_width = 0

        # while yielding rows __formula_values is updated with any formula values set on Expressions
        self.__formula_values = {}

        for name, (table, (row, col)) in list(self.__tables.items()):
            # get the resolved 2d data array from the table
            #
            # expressions with no explicit table will use None when calling
            # get_table/get_table_pos, which should return the current table.
            #
            self.__tables[None] = (table, (row, col))
            data = table.get_data(workbook, row, col, self.__formula_values)
            del self.__tables[None]

            height, width = data.shape
            upper_left = (row, col)
            lower_right = (row + height - 1, col + width - 1)

            max_height = max(max_height, lower_right[0] + 1)
            max_width = max(max_width, lower_right[1] + 1)
            
            resolved_tables.append((name, data, upper_left, lower_right))

        for row, col in self.__values.keys():
            max_width = max(max_width, row+1)
            max_height = max(max_height, col+1)

        # Build the whole table up-front. Doing it row by row is too slow.
        table = [[None] * max_width for i in range(max_height)]
        for name, data, upper_left, lower_right in resolved_tables:
            for i, r in enumerate(range(upper_left[0], lower_right[0]+1)):
                for j, c in enumerate(range(upper_left[1], lower_right[1]+1)):
                    table[r][c] = data[i][j]

        for (r, c), value in self.__values.items():
            if isinstance(value, Value):
                value = value.value
            if isinstance(value, Expression):
                if value.has_value:
                    self.__formula_values[(r, c)] = value.value
                value = value.get_formula(workbook, r, c)
            table[r][c] = value

        for row in table:
            yield row
 
    def to_csv(self, writer):
        """
        Writes worksheet to a csv.writer object.
        :param writer: csv writer instance.
        """
        for row in self.iterrows():
            writer.writerow(row)

    def _get_column_widths(self):
        """return a dictionary of {col -> width}"""
        col_widths = {}
        for table, (row, col) in self.__tables.values():
            for colname, width in table.column_widths.items():
                ic = col + table.get_column_offset(colname)
                current_width = col_widths.setdefault(ic, width)
                col_widths[ic] = max(width, current_width)
        return col_widths

    def _get_all_styles(self):
        """
        return a dictionary of {(row, col) -> CellStyle}
        for all cells that use a non-default style.
        """
        _styles = {}
        def _get_style(bold=False, bg_col=None, border=None):
            if (bold, bg_col, border) not in _styles:
                _styles[(bold, bg_col, border)] = CellStyle(bold=bold,
                                                            bg_color=bg_col,
                                                            border=border)
            return _styles[(bold, bg_col, border)]

        ws_styles = {}
        for table, (row, col) in self.__tables.values():
            for r in range(row, row + table.header_height):
                for c in range(col, col + table.width):
                    if isinstance(table.header_style, dict):
                        col_name = table.dataframe.columns[c]
                        style = table.header_style.get(col_name, _get_style(bold=True))
                    else:
                        style = table.header_style or _get_style(bold=True)
                    ws_styles[(r, c)] = style

            for c in range(col, col + table.row_labels_width):
                for r in range(row + table.header_height, row + table.height):
                    if isinstance(table.index_style, dict):
                        row_name = table.dataframe.index[r]
                        style = table.index_style.get(row_name, _get_style(bold=True))
                    else:
                        style = table.index_style or _get_style(bold=True)
                    ws_styles[(r, c)] = style

            if table.style.stripe_colors or table.style.border:
                num_bg_cols = len(table.style.stripe_colors) if \
                    table.style.stripe_colors else 1
                bg_cols = table.style.stripe_colors if \
                    table.style.stripe_colors else None

                for i, row_offset in enumerate(range(table.header_height,
                                                     table.height)):
                    for c in range(col, col + table.width):
                        bg_col = bg_cols[i % num_bg_cols] if bg_cols else None
                        style = _get_style(bold=None, bg_col=bg_col, border=table.style.border)
                        if (row + row_offset, c) in ws_styles:
                            style = style + ws_styles[(row + row_offset, c)]
                        ws_styles[(row + row_offset, c)] = style

            for col_name, col_style in table.column_styles.items():
                try:
                    col_offset = table.get_column_offset(col_name)
                except KeyError:
                    continue
                for i, r in enumerate(range(row + table.header_height, row + table.height)):
                    style = col_style
                    if (r, col + col_offset) in ws_styles:
                        style = ws_styles[(r, col + col_offset)] + style
                    ws_styles[(r, col + col_offset)] = style

            for row_name, row_style in table.row_styles.items():
                try:
                    row_offset = table.get_row_offset(row_name)
                except KeyError:
                    continue
                for i, c in enumerate(range(col + table.row_labels_width, col + table.width)):
                    style = row_style
                    if (row + row_offset, c) in ws_styles:
                        style = ws_styles[(row + row_offset, c)] + style
                    ws_styles[(row + row_offset, c)] = style

            for (row_name, col_name), cell_style in table.cell_styles.items():
                try:
                    col_offset = table.get_column_offset(col_name)
                    row_offset = table.get_row_offset(row_name)
                except KeyError:
                    continue
                style = cell_style
                if (row + row_offset, col + col_offset) in ws_styles:
                    style = ws_styles[(row + row_offset, col + col_offset)] + style
                ws_styles[(row + row_offset, col + col_offset)] = style

        for (row, col), value in self.__values.items():
            if isinstance(value, Value):
                style = value.style
                if style:
                    if (row, col) in ws_styles:
                        style = style + ws_styles[(row, col)]
                    ws_styles[(row, col)] = style

        return ws_styles

    def to_excel(self,
                 workbook=None,
                 worksheet=None,
                 xl_app=None,
                 clear=True,
                 rename=True,
                 resize_columns=True):
        """
        Writes worksheet to an Excel Worksheet COM object.
        Requires :py:module:`pywin32` to be installed.

        :param workbook: xltable.Workbook this sheet belongs to.
        :param worksheet: Excel COM Worksheet instance to write to.
        :param xl_app: Excel COM Excel Application to write to.
        :param bool clear: If a worksheet is provided, clear worksheet before writing.
        :param bool rename: If a worksheet is provided, rename self to match the worksheet.
        :param bool resize_columns: Resize sheet columns after writing.
        """
        from win32com.client import Dispatch, constants, gencache

        if xl_app is None:
            if worksheet is not None:
                xl_app = worksheet.Parent.Application
            elif workbook is not None and hasattr(workbook.workbook_obj, "Application"):
                xl_app = workbook.workbook_obj.Application
            else:
                xl_app = Dispatch("Excel.Application")

        xl = xl_app = gencache.EnsureDispatch(xl_app)

        # Create a workbook if there isn't one already
        if not workbook:
            from .workbook import Workbook
            workbook = Workbook(worksheets=[self])

        if worksheet is None:
            # If there's no worksheet then call Workbook.to_excel which will create one
            return workbook.to_excel(xl_app=xl_app, resize_columns=resize_columns)

        if rename:
            self.__name = worksheet.Name

        # set manual calculation and turn off screen updating while we update the cells
        calculation = xl.Calculation
        screen_updating = xl.ScreenUpdating
        xl.Calculation = constants.xlCalculationManual
        xl.ScreenUpdating = False
        try:
            # clear the worksheet and reset the styles
            if clear:
                worksheet.Cells.ClearContents()
                worksheet.Cells.Font.Bold = False
                worksheet.Cells.Font.Size = 11
                worksheet.Cells.Font.Color = 0x000000
                worksheet.Cells.Interior.ColorIndex = 0
                worksheet.Cells.NumberFormat = "General"

            origin = worksheet.Range("A1")
            xl_cell = origin
            for row in self.iterrows(workbook):
                row = _to_pywintypes(row)

                # set the value and formulae to the excel range (it's much quicker to
                # write a row at a time and update the formula than it is it do it
                # cell by cell)
                if clear:
                    xl_row = worksheet.Range(xl_cell, xl_cell.Offset(1, len(row)))
                    xl_row.Value = row
                else:
                    for i, value in enumerate(row):
                        if value is not None:
                            xl_cell.Offset(1, 1 + i).Value = value

                for i, value in enumerate(row):
                    if isinstance(value, str) and value.startswith("="):
                        xl_cell.Offset(1, 1 + i).Formula = value
                
                # move to the next row
                xl_cell = xl_cell.Offset(2, 1)

            # set any array formulas
            for table, (row, col) in self.__tables.values():
                if isinstance(table, ArrayFormula):
                    data = table.get_data(workbook, row, col)
                    height, width = data.shape
                    upper_left = origin.Offset(row+1, col+1)
                    lower_right = origin.Offset(row + height, col + width)

                    xl_range = worksheet.Range(upper_left, lower_right)
                    xl_range.FormulaArray = table.formula.get_formula(workbook, row, col)

            # set any formatting
            for (row, col), style in self._get_all_styles().items():
                r = origin.Offset(1 + row, 1 + col)
                if style.bold:
                    r.Font.Bold = True
                if style.excel_number_format is not None:
                    r.NumberFormat = style.excel_number_format
                if style.size is not None:
                    r.Font.Size = style.size
                if style.text_color is not None:
                    r.Font.Color = _to_bgr(style.text_color)
                if style.bg_color is not None:
                    r.Interior.Color = _to_bgr(style.bg_color)
                if style.text_wrap or style.border:
                    raise Exception("text wrap and border not implemented")

            # add any charts
            for chart, (row, col) in self.__charts:
                top_left = origin.Offset(1 + row, 1 + col)
                xl_chart = worksheet.ChartObjects().Add(top_left.Left, top_left.Top, 360, 220).Chart
                xl_chart.ChartType = _to_excel_chart_type(chart.type, chart.subtype)
                if chart.title:
                    xl_chart.ChartTitle = chart.title
                for series in chart.iter_series(self, row, col):
                    xl_series = xl_chart.SeriesCollection().NewSeries()
                    xl_series.Values = "=%s!%s" % (self.name, series["values"].lstrip("="))
                    if "categories" in series:
                        xl_series.XValues = "=%s!%s" % (self.name, series["categories"].lstrip("="))
                    if "name" in series:
                        xl_series.Name = series["name"]

        finally:
            xl.ScreenUpdating = screen_updating
            xl.Calculation = calculation

        if resize_columns:
            try:
                worksheet.Cells.EntireColumn.AutoFit()
            except:
                pass

    def to_xlsx(self, filename=None, workbook=None):
        """
        Write worksheet to a .xlsx file using xlsxwriter.

        :param str filename: Filename to write to. If None no file is written.
        :param xltable.Workbook: Workbook this sheet belongs to. If None a new workbook
        will be created with this worksheet as the only sheet.
        :return: :py:class:`xlsxwriter.workbook.Workbook` instance.
        """
        from .workbook import Workbook
        if not workbook:
            workbook = Workbook(filename=filename)
            workbook.append(self)
            return workbook.to_xlsx()
        ws = workbook.add_xlsx_worksheet(self, self.name)

        _styles = {}
        def _get_xlsx_style(cell_style):
            """
            convert rb.excel style to xlsx writer style
            """
            style_args = (
                cell_style.bold,
                cell_style.excel_number_format,
                cell_style.text_color,
                cell_style.bg_color,
                cell_style.size,
                cell_style.text_wrap,
                cell_style.text_wrap,
                cell_style.border,
                cell_style.align,
                cell_style.valign
            )
            if (style_args) not in _styles:
                style = workbook.add_format()
                if cell_style.bold:
                    style.set_bold()
                if cell_style.excel_number_format is not None:
                    style.set_num_format(cell_style.excel_number_format)
                if cell_style.text_color is not None:
                    style.set_font_color("#%06x" % cell_style.text_color)
                if cell_style.bg_color is not None:
                    style.set_bg_color("#%06x" % cell_style.bg_color)
                if cell_style.size is not None:
                    style.set_font_size(cell_style.size)
                if cell_style.text_wrap:
                    style.set_text_wrap()
                if cell_style.border:
                    style.set_border(cell_style.border)
                if cell_style.align:
                    style.set_align(cell_style.align)
                if cell_style.valign:
                    style.set_valign(cell_style.valign)

                _styles[style_args] = style

            return _styles[style_args]

        # pre-compute the cells with non-default styles
        ws_styles = self._get_all_styles()
        ws_styles = {(r, c): _get_xlsx_style(s) for ((r, c), s) in ws_styles.items()}
        plain_style = _get_xlsx_style(CellStyle())

        # write the rows to the worksheet
        for ir, row in enumerate(self.iterrows(workbook)):
            for ic, cell in enumerate(row):
                style = ws_styles.get((ir, ic), plain_style)
                if isinstance(cell, str):
                    if cell.startswith("="):
                        formula_value = self.__formula_values.get((ir, ic), 0)
                        ws.write_formula(ir, ic, cell, style, value=formula_value)
                    elif cell.startswith("{="):
                        continue
                    else:
                        ws.write(ir, ic, cell, style)
                else:
                    if isinstance(cell, self._xlsx_unsupported_types):
                        ws.write(ir, ic, str(cell), style)
                    else:
                        try:
                            ws.write(ir, ic, cell, style)
                        except TypeError:
                            ws.write(ir, ic, str(cell), style)
                            unsupported_types = set(self._xlsx_unsupported_types)
                            unsupported_types.add(type(cell))
                            self.__class__._xlsx_unsupported_types = tuple(unsupported_types)

        # set any array formulas
        for table, (row, col) in self.__tables.values():
            if isinstance(table, ArrayFormula):
                style = ws_styles.get((row, col), plain_style)
                data = table.get_data(workbook, row, col)
                height, width = data.shape
                bottom, right = (row + height - 1, col + width -1)
                formula = table.formula.get_formula(workbook, row, col)
                ws.write_array_formula(row, col, bottom, right, formula, style, value=data[0][0])

                for y in range(height):
                    for x in range(width):
                        if y == 0 and x == 0:
                            continue
                        ir, ic = row + y, col + x
                        style = ws_styles.get((ir, ic), plain_style)
                        cell = data[y][x]
                        if isinstance(cell, str):
                            cell_str = cell.encode("ascii", "xmlcharrefreplace").decode("ascii")
                            ws.write_formula(ir, ic, cell_str, style)
                        else:
                            ws.write(ir, ic, cell, style)

        # set any non-default column widths
        for ic, width in self._get_column_widths().items():
            ws.set_column(ic, ic, width)

        # add any charts
        for chart, (row, col) in self.__charts:
            kwargs = {"type": chart.type}
            if chart.subtype:
                kwargs["subtype"] = chart.subtype
            xl_chart = workbook.workbook_obj.add_chart(kwargs)

            if chart.show_blanks:
                xl_chart.show_blanks_as(chart.show_blanks)

            for series in chart.iter_series(workbook, row, col):
                # xlsxwriter expects the sheetname in the formula
                values = series.get("values")
                if isinstance(values, str) and values.startswith("=") and "!" not in values:
                    series["values"] = "='%s'!%s" % (self.name, values.lstrip("="))
                    
                categories = series.get("categories")
                if isinstance(categories, str) and categories.startswith("=") and "!" not in categories:
                    series["categories"] = "='%s'!%s" % (self.name, categories.lstrip("="))

                xl_chart.add_series(series)

            xl_chart.set_size({"width": chart.width, "height": chart.height})

            if chart.title:
                xl_chart.set_title({"name": chart.title})

            if chart.legend_position:
                xl_chart.set_legend({"position": chart.legend_position})

            if chart.x_axis:
                xl_chart.set_x_axis(chart.x_axis)

            if chart.y_axis:
                xl_chart.set_y_axis(chart.y_axis)

            ws.insert_chart(row, col, xl_chart)

        # add any groups
        for tables, collapsed in self.__groups:
            min_row, max_row = 1000000, -1

            for table, (row, col) in self.__tables.values():
                if table in tables:
                    min_row = min(min_row, row)
                    max_row = max(max_row, row + table.height)
            for i in range(min_row, max_row+1):
                ws.set_row(i, None, None, {'level': 1, 'hidden': collapsed})

        if filename:
            workbook.close()
        return workbook


def _to_bgr(rgb):
    """excel expects colors as BGR instead of the usual RGB"""
    if rgb is None:
        return None
    return ((rgb >> 16) & 0xff) + (rgb & 0xff00) + ((rgb & 0xff) << 16)


def _to_pywintypes(row):
    """convert values in a row to types accepted by excel"""
    def _pywintype(x):
        if isinstance(x, dt.date):
            return dt.datetime(x.year, x.month, x.day, tzinfo=dt.timezone.utc)

        elif isinstance(x, (dt.datetime, pa.Timestamp)):
            if x.tzinfo is None:
                return x.replace(tzinfo=dt.timezone.utc)

        elif isinstance(x, str):
            if re.match("^\d{4}-\d{2}-\d{2}$", x):
                return "'" + x
            return x

        elif isinstance(x, np.integer):
            return int(x)

        elif isinstance(x, np.floating):
            return float(x)

        elif x is not None and not isinstance(x, (str, int, float, bool)):
            return str(x)

        return x

    return [_pywintype(x) for x in row]


def _to_excel_chart_type(type, subtype):
    from win32com.client import constants
    return {
        "area": {
            None: constants.xlArea,
            "stacked": constants.xlAreaStacked,
            "percent_stacked": constants.xlAreaStacked100,
        },
        "bar": {
            None: constants.xlBar,
            "stacked": constants.xlBarStacked,
            "percent_stacked": constants.xlBarStacked100,
        },
        "column": {
            "stacked": constants.xlColumnStacked,
            "percent_stacked": constants.xlColumnStacked100,
        },
        "line": {
            None: constants.xlLine,
        },
        "scatter": {
            None: constants.xlXYScatter,
            "straight_with_markers": constants.xlXYScatterLines,
            "straight": constants.xlXYScatterLinesNoMarkers,
            "smooth_with_markers": constants.xlXYScatterSmooth,
            "smooth": constants.xlXYScatterSmoothNoMarkers,
        },
        "stock": {
            None: constants.xlStockHLC,
        },
        "radar": {
            None: constants.xlRadar,
            "with_markers": constants.xlRadarMarkers,
            "filled": constants.xlRadarFilled,
        },
    }[type][subtype]
