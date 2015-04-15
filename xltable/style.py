"""
Style objects are used when constructing Table instances
to direct how the tables and cells in the tables will be
written to Excel.
"""


class TableStyle(object):
    """
    Style to be applied to a table.

    :param tuple stripe_colors: Background cell colors to use.
    """
    def __init__(self, stripe_colors=(0xEAF1FA, 0xFFFFFF), border=None):
        self.stripe_colors = stripe_colors
        self.border = border


class CellStyle(object):
    """
    Style to be applied to a cell or range of cells.

    :param bool is_percentage: True if the cell value is a percentage.
    :param int decimal_places: Number of decimal places to display the cell value to.
    :param str date_format: Format to use for date values (use Python date format, eg '%Y-%m-%d').
    :param bool thousands_sep: True to display numbers with thousand separator.
    :param bool bold: True to make cells bold.
    :param int size: Text size, or use one of the string size aliases x-small, small, normal, large, x-large
    or xx-large.
    :param int text_color: Text color as an RGB value, eg 0xFF0000 for red.
    :param int bg_color: Background color as an RGB value, eg 0xFF0000 for red.
    """
    _sizes = {
        "x-small": 6,
        "small": 8,
        "normal": 11,
        "large": 16,
        "x-large": 20,
        "xx-large": 24
    }

    def __init__(self,
                 is_percentage=False,
                 decimal_places=None,
                 date_format=None,
                 thousands_sep=False,
                 bold=False,
                 size=None,
                 text_color=None,
                 bg_color=None,
                 text_wrap=False,
                 border=None,
                 align=None,
                 valign=None):
        self.is_percentage = is_percentage
        self.decimal_places = decimal_places
        self.date_format = date_format
        self.thousands_sep = thousands_sep
        self.bold = bold
        if isinstance(size, str):
            size = self._sizes[size]
        self.size = size
        self.text_color = text_color
        self.bg_color = bg_color
        self.text_wrap = text_wrap
        self.border = border
        self.align = align
        self.valign = valign

    @property
    def excel_number_format(self):
        number_format = "0"
        if self.thousands_sep:
            number_format = "#,#00"

        if self.decimal_places is not None:
            if self.decimal_places > 0:
                number_format = number_format + "." +  ("0" * self.decimal_places)

        if self.is_percentage:
            number_format = number_format + "%"

        if self.date_format is not None:
            number_format = self.date_format
            number_format = number_format.replace("%Y", "yyyy")
            number_format = number_format.replace("%m", "mm")
            number_format = number_format.replace("%d", "dd")
            number_format = number_format.replace("%H", "hh")
            number_format = number_format.replace("%M", "mm")
            number_format = number_format.replace("%S", "ss")

        if number_format == "0":
            return None
        return number_format
