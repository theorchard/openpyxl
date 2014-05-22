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
                 ht=None,
                 s=None,
                 hidden=False,
                 outlineLevel=0,
                 outline_level=None,
                 collapsed=False,
                 style=None,
                 visible=None,
                 height=None):
        if height is not None:
            ht = height
        self.ht = ht
        if style is not None:
            s = style
        self.s = s
        if visible is not None:
            hidden = not visible
        if outline_level is not None:
            outlineLevel = outlineLevel
        super(RowDimension, self).__init__(index, hidden, outlineLevel,
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

    __fields__ = Dimension.__fields__ + ('width', 'bestFit', 'customWidth')

    def __init__(self,
                 index='A',
                 width=None,
                 bestFit=False,
                 hidden=False,
                 outlineLevel=0,
                 outline_level=None,
                 collapsed=False,
                 style=None,
                 min=1,
                 max=1,
                 customWidth=False,
                 visible=None,
                 auto_size=None):
        self.width = width
        if width is not None:
            customWidth = True
        self.customWidth = customWidth
        self.style = style
        self.min = min
        self.max = max
        if visible is not None:
            hidden = not visible
        if auto_size is not None:
            bestFit = auto_size
        self.bestFit = bestFit
        if outline_level is not None:
            outlineLevel = outline_level
        super(ColumnDimension, self).__init__(index, hidden, outlineLevel,
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
