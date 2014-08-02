from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl


from .cell import Cell


class WriteOnlyCell(Cell):

    """
    Adapted for writing only
    """

    def __init__(self, worksheet=None, column='A', row=1, value=None):
        Cell.__init__(self, worksheet, column, row, value)

    @property
    def style(self):
        return self.parent.parent.shared_styles[self._style]

    @style.setter
    def style(self, new_style):
        self._style= self.parent.parent.shared_styles.add(new_style)
