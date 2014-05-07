from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

from openpyxl.compat import safe_string
from openpyxl.cell import get_column_letter, column_index_from_string
from openpyxl.descriptors import Integer, Float, Bool, Strict, String, Alias


class Dimension(Strict):
    """Information about the display properties of a row or column."""
    __fields__ = ('index',
                 'hidden',
                 'outlineLevel',
                 'collapsed',)

    __slots__ = __fields__
    index = Integer()
    hidden = Bool()
    outlineLevel = Integer(allow_none=True)
    outline_level = Alias('outlineLevel')
    collapsed = Bool()

    def __init__(self, index, hidden, outlineLevel,
                 collapsed, visible=True):
        self.index = index
        self.hidden = hidden
        self.outlineLevel = outlineLevel
        self.collapsed = collapsed

    def __iter__(self):
        for key in self.__fields__:
            value = getattr(self, key)
            if value:
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
                 collapsed=False,
                 visible=None):
        self.height = height
        if visible is not None:
            hidden = not visible
        super(RowDimension, self).__init__(index, hidden, outline_level,
                                           collapsed)


class ColumnDimension(Dimension):
    """Information about the display properties of a column."""

    width = Float(allow_none=True)
    bestFit = Bool()
    auto_size = Alias('bestFit')
    collapsed = Bool()
    index = String()

    __fields__ = Dimension.__fields__ + ('width', 'bestFit')

    def __init__(self,
                 index='A',
                 width=None,
                 auto_size=False,
                 hidden=False,
                 outline_level=0,
                 collapsed=False,
                 visible=None):
        self.width = width
        self.auto_size = auto_size
        if visible is not None:
            hidden = not visible
        super(ColumnDimension, self).__init__(index, hidden, outline_level,
                                              collapsed)


    @property
    def min(self):
        return column_index_from_string(self.index)

    @property
    def max(self):
        return self.min

    def __iter__(self):
        attrs = list(self.__fields__) + ['min', 'max']
        del attrs[0]
        for key in attrs:
            value = getattr(self, key)
            if value:
                yield key, safe_string(value)

    #@property
    #def col_label(self):
        #return get_column_letter(self.index)
