from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

from openpyxl.compat import safe_string
from openpyxl.cell import get_column_letter
from openpyxl.descriptors import Integer, Float, Bool, Strict, String, Alias


class Dimension(Strict):
    """Information about the display properties of a row or column."""
    __fields__ = ('index',
                 'hidden',
                 'outline_level',
                 'collapsed',)

    __slots__ = __fields__
    index = Integer()
    hidden = Bool()
    outline_level = Integer(allow_none=True)
    collapsed = Bool()

    def __init__(self, index, hidden, outline_level,
                 collapsed, visible=True):
        self.index = index
        self.hidden = hidden
        self.outline_level = outline_level
        self.collapsed = collapsed

    def __iter__(self):
        for key in self.__fields__:
            value = getattr(self, key)
            if value is not None:
                yield key, safe_string(value)

    @property
    def visible(self):
        return not self.hidden


class RowDimension(Dimension):
    """Information about the display properties of a row."""

    __fields__ = Dimension.__fields__ + ('ht',)
    ht = Float(allow_none=True)
    height = Alias('ht')

    def __init__(self,
                 index=0,
                 height=None,
                 hidden=False,
                 outline_level=0,
                 collapsed=False):
        super(RowDimension, self).__init__(index, hidden, outline_level,
                                           collapsed)
        self.height = height


class ColumnDimension(Dimension):
    """Information about the display properties of a column."""

    width = Float(allow_none=True)
    auto_size = Bool()
    collapsed = Bool()
    index = String()

    __fields__ = Dimension.__fields__ + ('width', 'auto_size')

    def __init__(self,
                 index='A',
                 width=None,
                 auto_size=False,
                 hidden=False,
                 outline_level=0,
                 collapsed=False):
        super(ColumnDimension, self).__init__(index, hidden, outline_level,
                                              collapsed)
        self.width = width
        self.auto_size = auto_size

    #@property
    #def col_label(self):
        #return get_column_letter(self.index)
