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

from openpyxl.descriptors import Float, Set

from .colors import WHITE
from .hashable import HashableObject
import warnings

from .descriptors import Color


FILL_NONE = 'none'
FILL_SOLID = 'solid'
FILL_GRADIENT_LINEAR = 'linear'
FILL_GRADIENT_PATH = 'path'
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

fills = (FILL_NONE, FILL_SOLID, FILL_GRADIENT_LINEAR, FILL_GRADIENT_PATH,
         FILL_PATTERN_DARKDOWN, FILL_PATTERN_DARKGRAY, FILL_PATTERN_DARKGRID,
         FILL_PATTERN_DARKHORIZONTAL, FILL_PATTERN_DARKTRELLIS, FILL_PATTERN_DARKUP,
         FILL_PATTERN_DARKVERTICAL, FILL_PATTERN_GRAY0625, FILL_PATTERN_GRAY125,
         FILL_PATTERN_LIGHTDOWN, FILL_PATTERN_LIGHTGRAY, FILL_PATTERN_LIGHTGRID,
         FILL_PATTERN_LIGHTHORIZONTAL, FILL_PATTERN_LIGHTTRELLIS,
         FILL_PATTERN_LIGHTUP, FILL_PATTERN_LIGHTVERTICAL, FILL_PATTERN_MEDIUMGRAY)


class Fill(HashableObject):
    """Area fill patterns for use in styles.
    Caution: if you do not specify a fill_type, other attributes will have
    no effect !"""

    __fields__ = ('fill_type',
                  'rotation',
                  'start_color',
                  'end_color')
    __check__ = {'start_color': Color,
                 'end_color': Color}
    fill_type = Set(values=fills)
    rotation = Float()
    start_color = Color()
    end_color = Color()

    def __init__(self, fill_type=FILL_NONE, rotation=0,
                 start_color=None, end_color=None):
        self.fill_type = fill_type
        self.rotation = rotation
        self.start_color = start_color
        self.end_color = end_color
