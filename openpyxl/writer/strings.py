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

"""Write the shared string table."""
from io import BytesIO

# package imports
from openpyxl.xml.functions import start_tag, end_tag, tag, XMLGenerator
from openpyxl.collections import IndexedList


def create_string_table(workbook):
    """Compile the string table for a workbook."""

    strings = set()
    for sheet in workbook.worksheets:
        for cell in sheet.get_cell_collection():
            if cell.data_type == cell.TYPE_STRING and cell._value is not None:
                strings.add(cell.value)
    return IndexedList(sorted(strings))


def write_string_table(string_table):
    """Write the string table xml."""
    temp_buffer = BytesIO()
    doc = XMLGenerator(out=temp_buffer)
    start_tag(doc, 'sst', {'xmlns':
            'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
            'uniqueCount': '%d' % len(string_table)})
    for key in string_table:
        start_tag(doc, 'si')
        if key.strip() != key:
            attr = {'xml:space': 'preserve'}
        else:
            attr = {}
        tag(doc, 't', attr, key)
        end_tag(doc, 'si')
    end_tag(doc, 'sst')
    string_table_xml = temp_buffer.getvalue()
    temp_buffer.close()
    return string_table_xml
