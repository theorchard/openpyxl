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

from openpyxl.descriptors import Float, Set, Sequence, Alias, Typed
from openpyxl.compat import safe_string

from .colors import WHITE, Color
from .hashable import HashableObject
import warnings


FILL_NONE = 'none'
FILL_SOLID = 'solid'
FILL_PATTERN_DARKDOWN = 'darkDown'
FILL_PATTERN_DARKGRAY = 'darkGray'
FILL_PATTERN_DARKGRID = 'darkGrid'
FILL_PATTERN_DARKHORIZONTAL = 'darkHorizontal'
FILL_PATTERN_DARKTRELLIS = 'darkTrellis'
FILL_PATTERN_DARKUP = 'darkUp'
FILL_PATTERN_DARKVERTICAL = 'darkVertical'
FILL_PATTERN_GRAY0625 = 'gray0625'
FILL_PATTERN_GRAY125 = 'gray125'
FILL_PATTERN_LIGHTDOWN = 'lightDown'
FILL_PATTERN_LIGHTGRAY = 'lightGray'
FILL_PATTERN_LIGHTGRID = 'lightGrid'
FILL_PATTERN_LIGHTHORIZONTAL = 'lightHorizontal'
FILL_PATTERN_LIGHTTRELLIS = 'lightTrellis'
FILL_PATTERN_LIGHTUP = 'lightUp'
FILL_PATTERN_LIGHTVERTICAL = 'lightVertical'
FILL_PATTERN_MEDIUMGRAY = 'mediumGray'

fills = (FILL_NONE, FILL_SOLID, FILL_PATTERN_DARKDOWN, FILL_PATTERN_DARKGRAY,
         FILL_PATTERN_DARKGRID, FILL_PATTERN_DARKHORIZONTAL, FILL_PATTERN_DARKTRELLIS,
         FILL_PATTERN_DARKUP, FILL_PATTERN_DARKVERTICAL, FILL_PATTERN_GRAY0625,
         FILL_PATTERN_GRAY125, FILL_PATTERN_LIGHTDOWN, FILL_PATTERN_LIGHTGRAY,
         FILL_PATTERN_LIGHTGRID, FILL_PATTERN_LIGHTHORIZONTAL,
         FILL_PATTERN_LIGHTTRELLIS, FILL_PATTERN_LIGHTUP, FILL_PATTERN_LIGHTVERTICAL,
         FILL_PATTERN_MEDIUMGRAY)


class Fill(HashableObject):

    """Base class"""

    pass


class PatternFill(Fill):
    """Area fill patterns for use in styles.
    Caution: if you do not specify a fill_type, other attributes will have
    no effect !"""

    spec = """18.8.32"""

    __fields__ = ('patternType',
                  'fgColor',
                  'bgColor')

    patternType = Set(values=fills)
    fill_type = Alias("patternType")
    fgColor = Typed(expected_type=Color)
    start_color = Alias("fgColor")
    bgColor = Typed(expected_type=Color)
    end_color = Alias("bgColor")

    def __init__(self, patternType=FILL_NONE, fgColor=Color(), bgColor=Color(),
                 fill_type=None, start_color=None, end_color=None):
        if fill_type is not None:
            patternType = fill_type
        self.patternType = patternType
        if start_color is not None:
            fgColor = start_color
        self.fgColor = fgColor
        if end_color is not None:
            bgColor = end_color
        self.bgColor = bgColor


class GradientFill(Fill):

    spec = """18.8.24"""

    __fields__ = ('fill_type', 'degree', 'left', 'right', 'top', 'bottom', 'stop')
    fill_type = Set(values=('linear', 'path'))
    type = Alias("fill_type")
    degree = Float()
    left = Float()
    right = Float()
    top = Float()
    bottom = Float()
    stop = Sequence(expected_type=Color)


    def __init__(self, fill_type="linear", degree=0, left=0, right=0, top=0,
                 bottom=0, stop=(), type=None):
        self.degree = degree
        self.left = left
        self.right = right
        self.top = top
        self.bottom = bottom
        self.stop = stop
        # cannot use type attribute but allow it is an argument (ie. when parsing)
        if type is not None:
            fill_type = type
        self.fill_type = fill_type

    def __iter__(self):
        """
        Dictionary interface for easier serialising.
        All values converted to strings
        """
        for key in ('type', 'degree', 'left', 'right', 'top', 'bottom'):
            value = getattr(self, key)
            if bool(value):
                yield key, safe_string(value)
