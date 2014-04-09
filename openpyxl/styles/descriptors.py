# Copyright (c) 2010-2014 openpyxl

"""
Package level descriptors
"""

from openpyxl.descriptors import Typed, Default
from .colors import Color


class Color(Typed):

    expected_type = Color

    def __init__(self, name=None, **kw):
        print kw
        if "defaults" not in kw:
            kw['defaults'] = {}
        super(Color, self).__init__(name, **kw)

    def __call__(self):
        return self.expected_type()
