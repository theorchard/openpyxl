# Copyright (c) 2010-2014 openpyxl

from io import BytesIO
from zipfile import ZipFile

import pytest

from openpyxl.xml.constants import ARC_WORKBOOK


@pytest.fixture()
def DummyArchive():
    body = BytesIO()
    archive = ZipFile(body, "w")
    return archive


def test_hidden_sheets(datadir, DummyArchive):
    from .. workbook import read_sheets

    datadir.chdir()
    archive = DummyArchive
    with open("hidden_sheets.xml") as src:
        archive.writestr(ARC_WORKBOOK, src.read())
        sheets = read_sheets(archive)
    assert list(sheets) == [
        ('rId1', 'Blatt1', None),
        ('rId2', 'Blatt2', 'hidden'),
        ('rId3', 'Blatt3', 'hidden')
                             ]
