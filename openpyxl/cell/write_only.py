from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl


from .cell import Cell


class WriteOnlyCell(Cell):

    """
    Adapted for writing only
    """

    @property
    def style(self):
        return self.parent.parent.shared_styles[self._style_id]

    @style.setter
    def style(self, new_style):
        self._style_id = self.parent.parent.shared_styles.add(new_style)

