from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

from copy import deepcopy

from openpyxl.descriptors import Typed

from .alignment import Alignment
from .borders import Borders, Border
from .colors import Color
from .fills import PatternFill, GradientFill, Fill
from .fonts import Font
from .hashable import HashableObject
from .numbers import NumberFormat, is_date_format, is_builtin
from .protection import Protection


class Style(HashableObject):
    """Style object containing all formatting details."""
    __fields__ = ('font',
                  'fill',
                  'borders',
                  'alignment',
                  'number_format',
                  'protection')
    __base__ = True

    font = Typed(expected_type=Font)
    fill = Typed(expected_type=Fill)
    borders = Typed(expected_type=Borders)
    alignment = Typed(expected_type=Alignment)
    number_format = Typed(expected_type=NumberFormat)
    protection = Typed(expected_type=Protection)

    def __init__(self, font=Font(), fill=PatternFill(), borders=Borders(),
                 alignment=Alignment(), number_format=NumberFormat(),
                 protection=Protection()):
        self.font = font
        self.fill = fill
        self.borders = borders
        self.alignment = alignment
        self.number_format = number_format
        self.protection = protection

DEFAULTS = Style()
