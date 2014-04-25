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

"""Imports for the openpyxl package."""
import warnings

TEST_LXML = False
SYSTEM_LXML = False

try:
    from .tests import LXML
    TEST_LXML = LXML
except  ImportError:
    try:
        from lxml.etree import LXML_VERSION
        SYSTEM_LXML = LXML_VERSION >= (3, 3, 1, 0)
        if SYSTEM_LXML is False:
            warnings.warn("The installed version of lxml is too old to be used with openpyxl")
    except ImportError:
        SYSTEM_LXML = False

# lxml is going to be used if and *only if* the version is correct on the sytem
# and if it was not disabled using environment variables
LXML = (TEST_LXML and SYSTEM_LXML)


from openpyxl.workbook import Workbook
from openpyxl.reader.excel import load_workbook

__author__ = 'Eric Gazoni and contributors'

__version__ = '1.9.0' # major.minor.patch
