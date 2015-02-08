from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl


from openpyxl.compat import unicode

from openpyxl.cell import Cell
from openpyxl.utils.datetime  import from_excel
from openpyxl.styles import is_date_format, Style
from openpyxl.styles.numbers import BUILTIN_FORMATS
from openpyxl.styles.styleable import StyleableObject


class ReadOnlyCell(StyleableObject):

    __slots__ = ('sheet', 'row', 'column', '_value', 'data_type', '_style_id',
                 '_font_id', '_border_id', '_fill_id', '_alignment_id', '_protection_id',
                 '_number_format_id')

    def __init__(self, sheet, row, column, value, data_type=Cell.TYPE_NULL, style_id=None):
        super(ReadOnlyCell, self).__init__(sheet=sheet)
        self._value = None
        self.row = row
        self.column = column
        self.data_type = data_type
        self.sheet = sheet
        self.value = value
        self._style_id = style_id

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
        return self.data_type == Cell.TYPE_NUMERIC and is_date_format(self.number_format)

    @property
    def number_format(self):
        if not self.style_id:
            return
        self._number_format_id = self.style_id.number_format
        return super(ReadOnlyCell, self).number_format

    @property
    def style_id(self):
        if not self._style_id:
            return
        return self.parent.parent._cell_styles[self._style_id]

    @property
    def font(self):
        self._font_id = self.style_id.font
        return super(ReadOnlyCell, self).font

    @property
    def fill(self):
        self._fill_id = self.style_id.fill
        return super(ReadOnlyCell, self).fill

    @property
    def border(self):
        self._border_id = self.style_id.border
        return super(ReadOnlyCell, self).border

    @property
    def alignment(self):
        self._alignment_id = self.style_id.alignment
        return super(ReadOnlyCell, self).alignment

    @property
    def protection(self):
        self._protection_id = self.style_id.protection
        return super(ReadOnlyCell, self).protection


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

    @value.setter
    def value(self, value):
        if self._value is not None:
            raise AttributeError("Cell is read only")
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
        wb = self.parent.parent
        font = wb._fonts[self.style_id.font]
        fill = wb._fills[self.style_id.fill]
        alignment = wb._alignments[self.style_id.alignment]
        border = wb._borders[self.style_id.border]
        protection = wb._protections[self.style_id.protection]

        return Style(font=font, alignment=alignment, fill=fill,
                     number_format=self.number_format, border=border, protection=protection)


EMPTY_CELL = ReadOnlyCell(None, None, None, None)
