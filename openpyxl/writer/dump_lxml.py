from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

from lxml.etree import xmlfile, Element, SubElement, tounicode

from . dump_worksheet import DumpWorksheet, DESCRIPTORS_CACHE_SIZE
from . lxml_worksheet import write_format, write_sheetviews

from openpyxl.xml.constants import SHEET_MAIN_NS


class LXMLWorksheet(DumpWorksheet):


    __saved = False


    def get_temporary_file(self, filename):
        if self.__saved:
            raise WorkbookAlreadySaved('this workbook has already been saved '
                    'and cannot be modified or saved anymore.')

        if filename in self._descriptors_cache:
            fobj = self._descriptors_cache[filename]
            # re-insert the value so it does not get evicted
            # from cache soon
            del self._descriptors_cache[filename]
            self._descriptors_cache[filename] = fobj
        else:
            fobj = open(filename, 'rb+')
            self._descriptors_cache[filename] = fobj
            if len(self._descriptors_cache) > DESCRIPTORS_CACHE_SIZE:
                filename, fileobj = self._descriptors_cache.popitem(last=False)
                fileobj.close()
        return fobj

    def write_header(self):
        NSMAP = {None : SHEET_MAIN_NS}
        fobj = self.get_temporary_file(self._fileobj_header_name)

        with xmlfile(fobj) as xf:
            xf.element("worksheet", nsmap=NSMAP)
            pr = Element('sheetPr')
            SubElement(pr, 'outlinePr',
                       {'summaryBelow':
                        '%d' %  (self.show_summary_below),
                        'summaryRight': '%d' % (self.show_summary_right)})
            if self.page_setup.fitToPage:
                SubElement(pr, 'pageSetUpPr', {'fitToPage': '1'})
            xf.write(pr)
            del pr

            #write_sheetviews(xf, self)
            #write_format(xf, self)


    def close_content(self):
        pass

    def _get_content_generator(self):
        pass

    def append(self, row):
        pass


