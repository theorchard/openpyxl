from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

"""Write the shared string table."""
from io import BytesIO

# package imports
from openpyxl.xml.constants import SHEET_MAIN_NS
from openpyxl.xml.functions import Element, xmlfile, SubElement

def write_string_table(string_table):
    """Write the string table xml."""
    out = BytesIO()
    NSMAP = {None : SHEET_MAIN_NS}

    with xmlfile(out) as xf:
        with xf.element("sst", nsmap=NSMAP, uniqueCount="%d" % len(string_table)):

            for key in string_table:
                el = Element('si')
                if key.strip() != key:
                    el.set('xml:space', 'preserve')
                text = SubElement(el, 't')
                text.text = key
                xf.write(el)

    return  out.getvalue()
