"""
Table class for building structured worksheets from multiple blocks of
related data.
"""
from .expression import Expression
from .style import TableStyle, CellStyle
from functools import partial
import pandas as pa


class Value(object):
    """
    Value wrapper that can be used in a table to add a style.

    :param value: Value that will be written to the cell.
    :param xltable.CellStyle: Style to be applied to the cell.
    """
    def __init__(self, value, style=None):
        self.value = value
        self.style = style


class Table(object):
    """
    Represents of table of data to be written to Excel, and
    may include :py:class:`xltable.Expression`s that will be converted into Excel
    formulas when the table's position is fixed.

    :param str name: Name of the table so it can be referenced by other tables and charts.
    :param pandas.DataFrame dataframe: Dataframe containing the data for the table.
    :param bool include_columns: Include the column names when outputting.
    :param bool include_index: Include the index when outputting.
    :param xltable.TableStyle style: Table style, or one of the named styles 'default' or 'plain'.
    :param xltable.CellStyle column_styles: Dictionary of column names to styles or named styles.
    :param float column_widths: Dictionary of column names to widths.
    :param xltable.CellStyle header_style: Style or named style to use for the cells in the header row.
    :param xltable.CellStyle index_style: Style or named style to use for the cells in the index column.

    Named table styles:
        - default: blue stripes
        - plain: no style

    Named cell styles:
        - pct: pecentage with two decimal places.
        - iso-date: date in YYYY-MM-DD format.
        - 2dp: two decimal places.
        - 2dpc: thousand separated number to two decimal places.
    """
    _styles = {
        "default": TableStyle(),
        "plain": TableStyle(stripe_colors=None)
    }

    _named_styles = {
        "pct": CellStyle(is_percentage=True, decimal_places=2),
        "iso-date": CellStyle(date_format="%Y-%m-%d"),
        "2dp": CellStyle(decimal_places=2),
        "2dpc": CellStyle(decimal_places=2, thousands_sep=True),
    }

    def __init__(self,
                 name,
                 dataframe,
                 include_columns=True,
                 include_index=False,
                 style="default",
                 column_styles={},
                 column_widths={},
                 row_styles={},
                 header_style=None,
                 index_style=None):
        self.__name = name
        self.__df = dataframe
        self.__position = None
        self.__include_columns = include_columns
        self.__include_index = include_index
        self.__column_widths = column_widths

        if isinstance(style, TableStyle):
            self.__style = style
        else:
            self.__style = self._styles.get(style)

        self.__col_styles = {}
        for col, style in column_styles.items():
            if isinstance(style, CellStyle):
                self.__col_styles[col] = style
            else:
                self.__col_styles[col] = self._named_styles[style]

        self.__row_styles = {}
        for row, style in row_styles.items():
            if isinstance(style, CellStyle):
                self.__row_styles[row] = style
            else:
                self.__row_styles[row] = self._named_styles[style]

        self.header_style = header_style
        self.index_style = index_style

    @property
    def name(self):
        return self.__name

    @property
    def dataframe(self):
        return self.__df

    @property
    def style(self):
        return self.__style

    @property
    def column_styles(self):
        return self.__col_styles

    @property
    def row_styles(self):
        return self.__row_styles

    @property
    def column_widths(self):
        return self.__column_widths

    @property
    def cell_styles(self):
        """dict of {(row name, col name): style}"""
        styles = {}
        for colname, col in self.dataframe.items():
            for rowname, value in col.items():
                if isinstance(value, Value) and value.style is not None:
                    style = value.style
                    if not isinstance(style, CellStyle):
                        style = self._named_styles[style]
                    styles[(rowname, colname)] = style
        return styles

    @property
    def width(self):
        return len(self.dataframe.columns) + self.row_labels_width

    @property
    def height(self):
        return len(self.dataframe.index) + self.header_height

    @property
    def header_height(self):
        if self.__include_columns:
            if isinstance(self.dataframe.columns, pa.MultiIndex):
                return len(self.dataframe.columns.names)
            return 1
        return 0

    @property
    def row_labels_width(self):
        if self.__include_index:
            if isinstance(self.dataframe.index, pa.MultiIndex):
                return len(self.dataframe.index.names)
            return 1
        return 0

    def get_column_offset(self, col):
        try:
            offset = self.dataframe.columns.get_loc(col)
        except KeyError:
            raise KeyError("Column '%s' not found in table %s" % (col, self.name))
        offset += self.row_labels_width
        return offset

    def get_index_offset(self):
        if self.__include_index:
            return 0
        raise KeyError("Table '%s' has no index" % self.name)

    def get_row_offset(self, row):
        try:
            offset = self.dataframe.index.get_loc(row)
        except KeyError:
            raise KeyError("Row '%s' not found in table %s" % (row, self.name))
        offset += self.header_height
        return offset

    def get_data(self, workbook, row, col, formula_values={}):
        """
        :return: 2d numpy array for this table with any formulas resolved to the final
        excel formula.
        :param xltable.Workbook workbook: Workbook the table has been added to.
        :param int row: Row where the table will start in the sheet (used for resolving formulas).
        :param int col: Column where the table will start in the sheet (used for resolving formulas).
        :param formula_values: dict to add pre-calculated formula values to (keyed by row, col).
        """
        if workbook:
            prev_table = workbook.active_table
            workbook.active_table = self
        try:
            df = self.dataframe.copy()

            # create a mask for elements that are expressions
            mask_df = df.applymap(lambda x: isinstance(x, Expression))

            # resolve any expressions if there are any
            if mask_df.any().any():
                # create a dataframe for indexing both into the dataframe and with the column and
                # row numbers.
                idx = [[(r, c) for r in range(len((df.index)))] for c in range(len((df.columns)))]
                index_df = pa.DataFrame(dict(zip(df.columns, idx)), columns=df.columns, index=df.index)

                # convert everything to objects so mask setting works
                df = df.astype(object)

                col_offset = self.row_labels_width
                row_offset = self.header_height

                # resolve all elements and set back into the main dataframe
                def get_formula(df, element):
                    if pa.isnull(element):
                        return element
                    r, c = element
                    expr = df.iget_value(r, c)
                    r += row_offset
                    c += col_offset
                    if expr.has_value:
                        formula_values[(r + row, c + col)] = expr.value
                    return expr.get_formula(workbook, r, c)

                df[mask_df] = index_df[mask_df].applymap(partial(get_formula, df))

            # replace any Cell instances with their value
            if df.applymap(lambda x: isinstance(x, Value)).any().any():
                df = df.applymap(lambda x: x.value if isinstance(x, Value) else x)

            # add the index and or columns to the values part of the dataframe
            if self.__include_index or self.__include_columns:
                index = df.index

                if self.__include_columns:
                    # add the index names to the top of the index to create a new row for the column headers
                    if isinstance(index, pa.MultiIndex):
                        index_names = tuple((x or "" for x in df.index.names))
                        i = 1
                        while index_names in df.index:
                            index_names = tuple(("%s_%d" % (x or "", i) for x in df.index.names))
                            i += 1
                        index_tuples = [index_names] + list(df.index.astype(object))
                        if isinstance(df.columns, pa.MultiIndex):
                            blank_tuple = tuple([None] * len(df.index.names))
                            index_tuples = ([blank_tuple] * (len(df.columns.levels) - 1)) + index_tuples
                        index = pa.MultiIndex.from_tuples(index_tuples)
                    else:
                        index_name = df.index.name
                        i = 1
                        while index_name in df.index:
                            index_name = "%s_%d" % (df.index.name, i)
                            i += 1
                        index = [index_name] + list(df.index.astype(object))
                        if isinstance(df.columns, pa.MultiIndex):
                            index = ([None] * (len(df.columns.levels) - 1)) + index

                columns = df.columns
                if self.__include_index:
                    # add the column names to the left of the columns to create a new row for the index headers
                    if isinstance(columns, pa.MultiIndex):
                        columns_names = tuple((x or "" for x in df.columns.names))
                        i = 1
                        while columns_names in df.columns:
                            columns_names = tuple(("%s_%d" % (x or "", i) for x in df.columns.names))
                            i += 1
                        column_tuples = [columns_names] + list(df.columns.astype(object))
                        if isinstance(df.index, pa.MultiIndex):
                            blank_tuple = tuple([None] * len(df.columns.names))
                            column_tuples = ([blank_tuple] * (len(df.index.levels) - 1)) + column_tuples
                        columns = pa.MultiIndex.from_tuples(column_tuples)
                    else:
                        columns_name = df.columns.name or ""
                        i = 1
                        while columns_name in df.columns:
                            columns_name = "%s_%d" % (df.columns.name, i)
                            i += 1
                        columns = [columns_name] + list(df.columns.astype(object))
                        if isinstance(df.index, pa.MultiIndex):
                            columns = ([None] * (len(df.index.levels) - 1)) + columns

                df = df.reindex(index=index, columns=columns).astype(object)

                if self.__include_columns:
                    if isinstance(df.columns, pa.MultiIndex):
                        for i in range(len(df.columns.levels)):
                            df.iloc[i, :] = [c[i] for c in df.columns.values]
                    else:
                        df.iloc[0, :] = df.columns

                if self.__include_index:
                    if isinstance(df.index, pa.MultiIndex):
                        for i in range(len(df.index.levels)):
                            df.iloc[:, i] = [c[i] for c in df.index.values]
                    else:
                        df.iloc[:, 0] = df.index

            # return the values as an np array
            return df.values

        finally:
            if workbook:
                workbook.active_table = prev_table


class ArrayFormula(Table):
    """
    Represents an array formula to be written to Excel.

    Subclass of :py:class:`xltable.Table`.

    :param str name: Name of table so it can be referenced by other tables and charts.
    :param xltable.Formula formula: Array formula.
    :param int width: Number of columns.
    :param int height: Number of row.
    :param pandas.DataFrame value: Precalculated formula result to save in the workbook.
    :param bool include_columns: Include the column names when outputting `value`.
    :param bool include_index: Include the index when outputting `value`.
    :param xltable.TableStyle style: Table style, or one of the named styles 'default' or 'plain'.
    :param xltable.CellStyle column_styles: Dictionary of column names to styles or named styles.
    :param float column_widths: Dictionary of column names to widths.
    """

    def __init__(self,
                 name,
                 formula,
                 width,
                 height,
                 value=None,
                 include_columns=False,
                 include_index=False,
                 style="default",
                 column_styles={},
                 column_widths={}):
        self.__formula = formula
        df = pa.DataFrame({c: [None] * height for c in range(width)})
        if value:
            df = value
        self.value = value
        super(ArrayFormula, self).__init__(name,
                                           dataframe=df,
                                           include_columns=include_columns,
                                           include_index=include_index,
                                           style=style,
                                           column_styles=column_styles,
                                           column_widths=column_widths)

    @property
    def formula(self):
        return self.__formula

    def get_data(self, workbook, row, col):
        if not self.value:
            self.dataframe[:] = "{%s}" % self.formula.get_formula(workbook, row, col)
        return super(ArrayFormula, self).get_data(workbook, row, col)
