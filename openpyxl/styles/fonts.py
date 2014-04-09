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

from openpyxl.descriptors import Strict, Float, Set, Bool, String, Typed, Integer
from .hashable import HashableObject
from .descriptors import Color


class Font(HashableObject):
    """Font options used in styles."""
    UNDERLINE_NONE = 'none'
    UNDERLINE_DOUBLE = 'double'
    UNDERLINE_DOUBLE_ACCOUNTING = 'doubleAccounting'
    UNDERLINE_SINGLE = 'single'
    UNDERLINE_SINGLE_ACCOUNTING = 'singleAccounting'

    name = String()
    size = Integer()
    bold = Bool()
    italic = Bool()
    superscript = Bool()
    subscript = Bool()
    underline = Set(values=set([UNDERLINE_DOUBLE, UNDERLINE_NONE,
                                UNDERLINE_DOUBLE_ACCOUNTING, UNDERLINE_SINGLE,
                                UNDERLINE_SINGLE_ACCOUNTING]))
    color = Color()

    __fields__ = ('name',
                  'size',
                  'bold',
                  'italic',
                  'superscript',
                  'subscript',
                  'underline',
                  'strikethrough',
                  'color')

    def __init__(self, name='Calibri', size=11, bold=False, italic=False,
                 superscript=False, subscript=False, underline=UNDERLINE_NONE,
                 strikethrough=False, color=color()):
        self.name = name
        self.size = size
        self.bold = bold
        self.italic = italic
        self.superscript = superscript
        self.subscript = subscript
        self.underline = underline
        self.strikethrough = strikethrough
        self.color = color
