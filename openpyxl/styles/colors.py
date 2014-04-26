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

from .hashable import HashableObject
from openpyxl.descriptors import String

# Default Color Index as per 18.8.27 of ECMA Part 4
COLOR_INDEX = ('00000000', '00FFFFFF', '00FF0000', '0000FF00', '000000FF',
               '00FFFF00', '00FF00FF', '0000FFFF', '00000000', '00FFFFFF',
               '00FF0000', '0000FF00', '000000FF', '00FFFF00', '00FF00FF', '0000FFFF',
               '00800000', '00008000', '00000080', '00808000', '00800080', '00008080',
               '00C0C0C0', '00808080', '009999FF', '00993366', '00FFFFCC', '00CCFFFF',
               '00660066', '00FF8080', '000066CC', '00CCCCFF', '00000080', '00FF00FF',
               '00FFFF00', '0000FFFF', '00800080', '00800000', '00008080', '000000FF',
               '0000CCFF', '00CCFFFF', '00CCFFCC', '00FFFF99', '0099CCFF', '00FF99CC',
               '00CC99FF', '00FFCC99', '003366FF', '0033CCCC', '0099CC00', '00FFCC00',
               '00FF9900', '00FF6600', '00666699', '00969696', '00003366', '00339966',
               '00003300', '00333300', '00993300', '00993366', '00333399', 'System Foreground', 'System Background')

BLACK = COLOR_INDEX[0]
WHITE = COLOR_INDEX[1]
RED = COLOR_INDEX[2]
DARKRED = COLOR_INDEX[8]
BLUE = COLOR_INDEX[4]
DARKBLUE = COLOR_INDEX[10]
GREEN = COLOR_INDEX[3]
DARKGREEN = COLOR_INDEX[9]
YELLOW = COLOR_INDEX[5]
DARKYELLOW = COLOR_INDEX[11]


class Color(HashableObject):
    """Named colors for use in styles."""

    index = String('_index')

    __fields__ = ('index',)

    def __init__(self, index=BLACK):
        self.index = index
