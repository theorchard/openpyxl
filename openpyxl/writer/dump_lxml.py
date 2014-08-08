from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

from lxml.etree import xmlfile, Element, SubElement, tounicode

from . dump_worksheet import DumpWorksheet, DESCRIPTORS_CACHE_SIZE
from . lxml_worksheet import write_format, write_sheetviews

from openpyxl.xml.constants import SHEET_MAIN_NS


class LXMLWorksheet(DumpWorksheet):

    __saved = False

    def write_header(self):
        NSMAP = {None : SHEET_MAIN_NS}

        with xmlfile(self.filename) as xf:
            with xf.element("worksheet", nsmap=NSMAP):
                pr = Element('sheetPr')
                SubElement(pr, 'outlinePr',
                           {'summaryBelow':
                            '%d' %  (self.show_summary_below),
                            'summaryRight': '%d' % (self.show_summary_right)})
                if self.page_setup.fitToPage:
                    SubElement(pr, 'pageSetUpPr', {'fitToPage': '1'})
                xf.write(pr)

                dim = Element('dimension', {'ref': 'A1:%s' % (self.get_dimensions())})
                xf.write(dim)

                write_sheetviews(xf, self)
                write_format(xf, self)


    def close_content(self):
        pass

    def _get_content_generator(self):
        pass

    def append(self, row):
        pass


