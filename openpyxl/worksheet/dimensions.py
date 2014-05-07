from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl


from openpyxl.cell import get_column_letter
from openpyxl.descriptors import Integer, Float, Bool, Strict, String


class Dimension(Strict):
    """Information about the display properties of a row or column."""
    __fields__ = ('index',
                 'visible',
                 'outline_level',
                 'collapsed',)

    __slots__ = __fields__
    index = Integer()
    visible = Bool()
    outline_level = Integer(allow_none=True)
    collapsed = Bool()

    def __init__(self, index=0, visible=True, outline_level=0,
                 collapsed=False):
        self.index = index
        self.visible = visible
        self.outline_level = outline_level
        self.collapsed = collapsed


class RowDimension(Dimension):
    """Information about the display properties of a row."""

    __fields__ = Dimension.__fields__ + ('height',)
    height = Float(allow_none=True)

    def __init__(self,
                 index=0,
                 height=None,
                 visible=True,
                 outline_level=0,
                 collapsed=False):
        super(RowDimension, self).__init__(index, visible, outline_level,
                                           collapsed)
        if height is not None:
            height = float(height)
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
                 visible=True,
                 outline_level=0,
                 collapsed=False):
        super(ColumnDimension, self).__init__(index, visible, outline_level,
                                              collapsed)
        self.width = width
        self.auto_size = auto_size

    #@property
    #def col_label(self):
        #return get_column_letter(self.index)
