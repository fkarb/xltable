"""
Chart objects reference data from Table instances and are written
to Excel worksheets as Excel charts.
"""
import datetime as dt


class Chart(object):
    """
    Chart objects reference data from Table instances and are written
    to Excel worksheets as Excel charts.

    :param str type: Chart type (see below).
    :param str subtype: Chart sub type (see below).
    :param str title: Chart title
    :param str legend_position: right (default), left, top, bottom or 'none' for no legend.
    :param int width: Chart width.
    :param int height: Chart height.

    Chart types and sub-types:
        - area:
            - stacked
            - percent_stacked
        - bar:
            - stacked
            - perecent_stacked
        - column:
            - stacked
            - perecent_stacked
        - line
        - scatter:
            - straight_with_markers
            - straight
            - smooth_with_markers
            - smooth
        - stock
        - radar:
            - with_markers
            - filled
    """

    def __init__(self,
                 type,
                 subtype=None,
                 title=None,
                 legend_position=None,
                 x_axis=None,
                 y_axis=None,
                 show_blanks=None,  # set to 'gap', 'zero' or 'span'
                 width=480,
                 height=288):
        self.type = type
        self.subtype = subtype
        self.title = title
        self.legend_position = legend_position
        self.x_axis = dict(x_axis) if x_axis else x_axis
        self.y_axis = dict(y_axis) if y_axis else y_axis
        self.show_blanks = show_blanks
        self.width = width
        self.height = height
        self.__series = []

        # convert dates in the axis args to serial dates
        for axis in (self.x_axis, self.y_axis):
            if axis:
                for key, value in list(axis.items()):
                    if isinstance(value, dt.date):
                        axis[key] = (value - dt.date(1900, 1, 1)).days + 2

    def add_series(self, values, **kwargs):
        """
        Adds a series to the chart.
        
        :param values: A :py:class:`xltable.Expression` object that evaluates to the data series.
        :param categories: A :py:class:`xltable.Expression` object that evaluates to the data series.
        :param name: Name to show in the legend for the series
        :param line: Line style, eg {'color': 'blue', 'width': 3.25} or {'none': True}
        :param marker: dict specifying how the markers should look, eg {type: square}.
        :param trendline: dict specifying how the trendline should be drawn, eg {type: linear}.
        """
        series = {"values": values}
        series.update(kwargs)
        self.__series.append(series)

    def iter_series(self, workbook, row, col):
        """
        Yield series dictionaries with values resolved to the final excel formulas.
        """
        for series in self.__series:
            series = dict(series)
            series["values"] = series["values"].get_formula(workbook, row, col)
            if "categories" in series:
                series["categories"] = series["categories"].get_formula(workbook, row, col)
            yield series
