from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl


from openpyxl.descriptors import (
    Strict,
    Typed,
)
from .fills import PatternFill, GradientFill, Fill
from . fonts import Font
from . borders import Border
from . alignment import Alignment
from . numbers import NumberFormatDescriptor
from . protection import Protection

from openpyxl.xml.constants import SHEET_MAIN_NS


class NamedStyle(object):

    tag = '{%s}cellXfs' % SHEET_MAIN_NS

    """
    Named and editable styles
    """

    font = Typed(expected_type=Font)
    fill = Typed(expected_type=Fill)
    border = Typed(expected_type=Border)
    alignment = Typed(expected_type=Alignment)
    number_format = NumberFormatDescriptor()
    protection = Typed(expected_type=Protection)

    def __init__(self,
                 name,
                 font=Font(),
                 fill=PatternFill(),
                 border=Border(),
                 alignment=Alignment(),
                 number_format=None,
                 protection=Protection()
                 ):
        self.name = name
        self.font = font
        self.fill = fill
        self.border = border
        self.alignment = alignment
        self.number_format = number_format
        self.protection = protection
