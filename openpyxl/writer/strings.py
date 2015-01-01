from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

"""Write the shared string table."""
from io import BytesIO

# package imports
from openpyxl.xml.functions import start_tag, end_tag, tag, XMLGenerator


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
