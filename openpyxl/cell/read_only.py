from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl


from openpyxl.compat import unicode

from openpyxl.cell import Cell
from openpyxl.utils.datetime  import from_excel
from openpyxl.styles import is_date_format, Style
from openpyxl.styles.numbers import BUILTIN_FORMATS


class ReadOnlyCell(object):

    __slots__ = ('sheet', 'row', 'column', '_value', 'data_type', '_style_id')

    def __init__(self, sheet, row, column, value, data_type=Cell.TYPE_NULL, style_id=None):
        self.row = row
        self.column = column
        self.data_type = data_type
        self.sheet = sheet
        self._set_value(value)
        self._set_style_id(style_id)

    def __eq__(self, other):
        for a in self.__slots__:
            if getattr(self, a) != getattr(other, a):
                return
        return True

    def __ne__(self, other):
        return not self.__eq__(other)

    @property
    def shared_strings(self):
        return self.sheet.shared_strings

    @property
    def base_date(self):
        return self.sheet.base_date

    @property
    def coordinate(self):
        if self.row is None or self.column is None:
            raise AttributeError("Empty cells have no coordinates")
        return "{1}{0}".format(self.row, self.column)

    @property
    def is_date(self):
        return self.data_type == Cell.TYPE_NUMERIC and is_date_format(self.number_format)

    @property
    def number_format(self):
        if self.style_id is None:
            return
        return self.style.number_format

    @property
    def style_id(self):
        return self._style_id

    def _set_style_id(self, value):
        if value is not None:
            value = int(value)
        self._style_id = value

    @property
    def internal_value(self):
        return self._value

    @property
    def value(self):
        if self._value is None:
            return
        if self.data_type == Cell.TYPE_BOOL:
            return self._value == '1'
        elif self.is_date:
            return from_excel(self._value, self.base_date)
        elif self.data_type in(Cell.TYPE_INLINE, Cell.TYPE_FORMULA_CACHE_STRING):
            return unicode(self._value)
        elif self.data_type == Cell.TYPE_STRING:
            return unicode(self.shared_strings[int(self._value)])
        return self._value

    def _set_value(self, value):
        if value is None:
            self.data_type = Cell.TYPE_NULL
        elif self.data_type == Cell.TYPE_NUMERIC:
            try:
                value = int(value)
            except ValueError:
                value = float(value)
        self._value = value

    @property
    def style(self):
        if self.style_id is None:
            return Style()
        wb = self.sheet.parent
        styles_id = wb._cell_styles[self.style_id]
        font = wb._fonts[styles_id.font]
        fill = wb._fills[styles_id.fill]
        alignment = wb._alignments[styles_id.alignment]
        border = wb._borders[styles_id.border]
        protection = wb._protections[styles_id.protection]
        if styles_id.number_format < 164:
            number_format = BUILTIN_FORMATS.get(styles_id.number_format, "General")
        else:
            number_format = wb._number_formats[styles_id.number_format - 164]

        return Style(font=font, alignment=alignment, fill=fill,
                     number_format=number_format, border=border, protection=protection)


EMPTY_CELL = ReadOnlyCell(None, None, None, None)
