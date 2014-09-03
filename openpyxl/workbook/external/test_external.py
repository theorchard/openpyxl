# Copyright (c) 2010-2014 openpyxl

from io import BytesIO
from zipfile import ZipFile

import pytest

from openpyxl.reader.workbook import read_rels
from openpyxl.xml.constants import (
    ARC_CONTENT_TYPES,
    ARC_WORKBOOK_RELS,
    PKG_REL_NS,
    REL_NS,
)


def test_read_external_ref(datadir):
    datadir.chdir()
    archive = ZipFile(BytesIO(), "w")
    with open("[Content_Types].xml") as src:
        archive.writestr(ARC_CONTENT_TYPES, src.read())
    with open("workbook.xml.rels") as src:
        archive.writestr(ARC_WORKBOOK_RELS, src.read())
    rels = read_rels(archive)
    for _, pth in rels:
        if pth['type'] == '%s/externalLink' % REL_NS:
            assert pth['path'] == 'xl/externalLinks/externalLink1.xml'


def test_read_external_link(datadir):
    from openpyxl.workbook.external import parse_books
    datadir.chdir()
    with open("externalLink1.xml.rels") as src:
        xml = src.read()
    books = tuple(parse_books(xml))
    assert books[0].Id == 'rId1'


def test_read_defined_names(datadir):
    from openpyxl.workbook.external import parse_names
    datadir.chdir()
    with open("externalLink1.xml") as src:
        xml = src.read()
    names = tuple(parse_names(xml))
    assert names[0].name == 'B2range'
    assert names[0].refersTo == "='Sheet1'!$A$1:$A$10"


def test_dict_external_book():
    from openpyxl.workbook.external import ExternalBook
    book = ExternalBook('rId1', "book1.xlsx")
    assert dict(book) == {'Id':'rId1', 'Target':'book1.xlsx',
                          'TargetMode':'External',
                          'Type':'http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLinkPath'}
