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

from openpyxl.descriptors import Set, Default, Typed

from .colors import Color
from .hashable import HashableObject
from .descriptors import Color
from .border import Border


class Border(Default):
    """Border descriptor"""

    expected_type = Border


DIAGONAL_NONE = 0
DIAGONAL_UP = 1
DIAGONAL_DOWN = 2
DIAGONAL_BOTH = 3
diagonals = (DIAGONAL_NONE, DIAGONAL_UP, DIAGONAL_DOWN, DIAGONAL_BOTH)


class Borders(HashableObject):
    """Border positioning for use in styles."""


    __fields__ = ('left',
                  'right',
                  'top',
                  'bottom',
                  'diagonal',
                  'diagonal_direction',
                  'all_borders',
                  'outline',
                  'inside',
                  'vertical',
                  'horizontal')

    left = Border()
    right = Border()
    top = Border()
    bottom = Border()
    diagonal = Border()
    diagonal_direction = Set(values=diagonals)
    all_borders = Border()
    outline = Border()
    inside = Border()
    vertical = Border()
    horizontal = Border()

    def __init__(self, left=left(), right=right(), top=top(),
                 bottom=bottom(), diagonal=diagonal(), diagonal_direction=DIAGONAL_NONE,
                 all_borders=all_borders(), outline=outline(), inside=inside(),
                 vertical=vertical(), horizontal=horizontal()):
        self.left = left
        self.right = right
        self.top = top
        self.bottom = bottom
        self.diagonal = diagonal
        self.all_borders = all_borders
        self.outline = outline
        self.inside = inside
        self.vertical = vertical
        self.horizontal = horizontal
        self.diagonal_direction = diagonal_direction
