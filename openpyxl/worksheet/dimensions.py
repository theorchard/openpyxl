from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

from openpyxl.compat import safe_string
from openpyxl.cell import get_column_letter, column_index_from_string
from openpyxl.descriptors import Integer, Float, Bool, Strict, String, Alias


class Base(Strict):
    # Base class for avoiding conflicts between descriptors and slots in Python 3
    __fields__ = ()
    __slots__ = __fields__


class Dimension(Base):
    """Information about the display properties of a row or column."""
    __fields__ = ('index',
                 'hidden',
                 'outlineLevel',
                 'collapsed',)

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
    s = Integer(allow_none=True)
    style = Alias('s')

    def __init__(self,
                 index=0,
                 height=None,
                 hidden=False,
                 outline_level=0,
                 collapsed=False,
                 style=None,
                 visible=None):
        self.height = height
        self.style = style
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
    style = Integer(allow_none=True)
    min = Integer()
    max = Integer()
    customWidth = Bool()

    __fields__ = Dimension.__fields__ + ('width', 'bestFit')

    def __init__(self,
                 index='A',
                 width=None,
                 bestFit=False,
                 hidden=False,
                 outline_level=0,
                 collapsed=False,
                 style=None,
                 min=1,
                 max=1,
                 customWidth=False,
                 visible=None,
                 auto_size=None):
        self.width = width
        self.style = style
        self.min = min
        self.max = max
        if visible is not None:
            hidden = not visible
        if auto_size is not None:
            bestFit = auto_size
        self.bestFit = auto_size
        super(ColumnDimension, self).__init__(index, hidden, outline_level,
                                              collapsed)

    def __iter__(self):
        for key in self.__fields__[1:]:
            value = getattr(self, key)
            if value:
                yield key, safe_string(value)

    #@property
    #def col_label(self):
        #return get_column_letter(self.index)

del Base
