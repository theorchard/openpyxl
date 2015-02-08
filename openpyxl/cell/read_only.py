from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl


from openpyxl.compat import unicode

from openpyxl.cell import Cell
from openpyxl.utils.datetime  import from_excel
from openpyxl.styles import is_date_format, Style
from openpyxl.styles.numbers import BUILTIN_FORMATS
from openpyxl.styles.styleable import StyleableObject


class ReadOnlyCell(StyleableObject):

    __slots__ = StyleableObject.__slots__ + ('sheet', 'row', 'column', '_value', 'data_type')

    def __init__(self, sheet, row, column, value, data_type='n', style_id=None):
        self.parent = sheet
        self._value = None
        self.row = row
        self.column = column
        self.data_type = data_type
        self.sheet = sheet
        self.value = value
        self._font_id = 0
        self._fill_id = 0
        self._border_id = 0
        self._alignment_id = 0
        self._protection_id = 0
        self._number_format_id = 0
        if style_id is not None:
            style = sheet.parent._cell_styles[style_id]
            self._font_id = style.font
            self._fill_id = style.fill
            self._border_id = style.border
            self._alignment_id = style.alignment
            self._protection_id = style.protection
            self._number_format_id = style.number_format

    def __eq__(self, other):
        for a in self.__slots__:
            if getattr(self, a) != getattr(other, a):
                return
        return True

    def __ne__(self, other):
        return not self.__eq__(other)

    @property
    def shared_strings(self):
        return self.parent.shared_strings

    @property
    def base_date(self):
        return self.parent.base_date

    @property
    def coordinate(self):
        if self.row is None or self.column is None:
            raise AttributeError("Empty cells have no coordinates")
        return "{1}{0}".format(self.row, self.column)

    @property
    def is_date(self):
        return self.data_type == 'n' and is_date_format(self.number_format)

    @property
    def internal_value(self):
        return self._value

    @property
    def value(self):
        if self._value is None:
            return
        if self.data_type == 'b':
            return self._value == '1'
        elif self.is_date:
            return from_excel(self._value, self.base_date)
        elif self.data_type in(Cell.TYPE_INLINE, Cell.TYPE_FORMULA_CACHE_STRING):
            return unicode(self._value)
        elif self.data_type == 's':
            return unicode(self.shared_strings[int(self._value)])
        return self._value

    @value.setter
    def value(self, value):
        if self._value is not None:
            raise AttributeError("Cell is read only")
        if value is None:
            self.data_type = 'n'
        elif self.data_type == 'n':
            try:
                value = int(value)
            except ValueError:
                value = float(value)
        self._value = value

    @property
    def style(self):
        wb = self.parent.parent
        font = wb._fonts[self.style_id.font]
        fill = wb._fills[self.style_id.fill]
        alignment = wb._alignments[self.style_id.alignment]
        border = wb._borders[self.style_id.border]
        protection = wb._protections[self.style_id.protection]

        return Style(font=font, alignment=alignment, fill=fill,
                     number_format=self.number_format, border=border, protection=protection)


EMPTY_CELL = ReadOnlyCell(None, None, None, None)
