from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in
# all copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
# THE SOFTWARE.
#
# @license: http://www.opensource.org/licenses/mit-license.php
# @author: see AUTHORS file

from openpyxl.compat import safe_string
from openpyxl.descriptors import Set, Alias

from .hashable import HashableObject
from .descriptors import Color


class Border(HashableObject):

    spec = """Actually to BorderPr 18.8.6"""

    """Border options for use in styles.
    Caution: if you do not specify a border_style, other attributes will
    have no effect !"""
    BORDER_NONE = None
    BORDER_DASHDOT = 'dashDot'
    BORDER_DASHDOTDOT = 'dashDotDot'
    BORDER_DASHED = 'dashed'
    BORDER_DOTTED = 'dotted'
    BORDER_DOUBLE = 'double'
    BORDER_HAIR = 'hair'
    BORDER_MEDIUM = 'medium'
    BORDER_MEDIUMDASHDOT = 'mediumDashDot'
    BORDER_MEDIUMDASHDOTDOT = 'mediumDashDotDot'
    BORDER_MEDIUMDASHED = 'mediumDashed'
    BORDER_SLANTDASHDOT = 'slantDashDot'
    BORDER_THICK = 'thick'
    BORDER_THIN = 'thin'

    __fields__ = ('style',
                  'color')

    color = Color(allow_none=True)
    style = Set(values=(BORDER_NONE, BORDER_DASHDOT, BORDER_DASHDOTDOT,
                        BORDER_DASHED, BORDER_DOTTED, BORDER_DOUBLE, BORDER_HAIR, BORDER_MEDIUM,
                        BORDER_MEDIUMDASHDOT, BORDER_MEDIUMDASHDOTDOT, BORDER_MEDIUMDASHED,
                        BORDER_SLANTDASHDOT, BORDER_THICK, BORDER_THIN))
    border_style = Alias('style')

    def __init__(self, style=None, color=None, border_style=None):
        if border_style is not None:
            style = border_style
        self.style = style
        self.color = color

    def __iter__(self):
        for key in ("style",):
            value = getattr(self, key)
            if value:
                yield key, safe_string(value)
