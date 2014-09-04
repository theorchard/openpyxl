# Copyright (c) 2010-2014 openpyxl

from io import BytesIO
from zipfile import ZipFile

import pytest

from openpyxl.xml.constants import (
    ARC_WORKBOOK,
    ARC_CONTENT_TYPES,
    ARC_WORKBOOK_RELS,
    REL_NS,
)


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


@pytest.mark.parametrize("excel_file, expected", [
    ("bug137.xlsx", [
        {'path': 'xl/worksheets/sheet1.xml', 'title': 'Sheet1', 'type':'%s/worksheet' % REL_NS}
        ]
     ),
    ("contains_chartsheets.xlsx", [
        {'path': 'xl/worksheets/sheet1.xml', 'title': 'data', 'type':'%s/worksheet' % REL_NS},
        {'path': 'xl/worksheets/sheet2.xml', 'title': 'moredata', 'type':'%s/worksheet' % REL_NS},
        ]),
    ("bug304.xlsx", [
    {'path': 'xl/worksheets/sheet3.xml', 'title': 'Sheet1', 'type':'%s/worksheet' % REL_NS},
    {'path': 'xl/worksheets/sheet2.xml', 'title': 'Sheet2', 'type':'%s/worksheet' % REL_NS},
    {'path': 'xl/worksheets/sheet.xml', 'title': 'Sheet3', 'type':'%s/worksheet' % REL_NS},
    ])
]
                         )
def test_detect_worksheets(datadir, excel_file, expected):
    from openpyxl.reader.excel import detect_worksheets

    datadir.chdir()
    archive = ZipFile(excel_file)
    assert list(detect_worksheets(archive)) == expected


@pytest.mark.parametrize("excel_file, expected", [
    ("bug137.xlsx", {
        "rId1": {'path': 'xl/chartsheets/sheet1.xml', 'type':'%s/chartsheet' % REL_NS},
        "rId2": {'path': 'xl/worksheets/sheet1.xml', 'type':'%s/worksheet' % REL_NS},
        "rId3": {'path': 'xl/theme/theme1.xml', 'type':'%s/theme' % REL_NS},
        "rId4": {'path': 'xl/styles.xml', 'type':'%s/styles' % REL_NS},
        "rId5": {'path': 'xl/sharedStrings.xml', 'type':'%s/sharedStrings' % REL_NS}
    }),
    ("bug304.xlsx", {
        'rId1': {'path': 'xl/worksheets/sheet3.xml', 'type':'%s/worksheet' % REL_NS},
        'rId2': {'path': 'xl/worksheets/sheet2.xml', 'type':'%s/worksheet' % REL_NS},
        'rId3': {'path': 'xl/worksheets/sheet.xml', 'type':'%s/worksheet' % REL_NS},
        'rId4': {'path': 'xl/theme/theme.xml', 'type':'%s/theme' % REL_NS},
        'rId5': {'path': 'xl/styles.xml', 'type':'%s/styles' % REL_NS},
        'rId6': {'path': '../customXml/item1.xml', 'type':'%s/customXml' % REL_NS},
        'rId7': {'path': '../customXml/item2.xml', 'type':'%s/customXml' % REL_NS},
        'rId8': {'path': '../customXml/item3.xml', 'type':'%s/customXml' % REL_NS}
    }),
]
                         )
def test_read_rels(datadir, excel_file, expected):
    from openpyxl.reader.workbook import read_rels

    datadir.chdir()
    archive = ZipFile(excel_file)
    assert dict(read_rels(archive)) == expected


@pytest.mark.parametrize("workbook_file, expected", [
    ("bug137_workbook.xml",
     [
         ("rId1", "Chart1", None),
         ("rId2", "Sheet1", None),
     ]
     ),
    ("bug304_workbook.xml",
     [
         ('rId1', 'Sheet1', None),
         ('rId2', 'Sheet2', None),
         ('rId3', 'Sheet3', None),
     ]
     )
])
def test_read_sheets(datadir, DummyArchive, workbook_file, expected):
    from openpyxl.reader.workbook import read_sheets

    datadir.chdir()
    archive = DummyArchive

    with open(workbook_file) as src:
        archive.writestr(ARC_WORKBOOK, src.read())
    assert list(read_sheets(archive)) == expected


def test_read_content_types(datadir, DummyArchive):
    from openpyxl.reader.workbook import read_content_types

    archive = DummyArchive
    datadir.chdir()
    with open("content_types.xml") as src:
        archive.writestr(ARC_CONTENT_TYPES, src.read())

    assert list(read_content_types(archive)) == [
    ('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml', '/xl/workbook.xml'),
    ('application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml', '/xl/worksheets/sheet1.xml'),
    ('application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml', '/xl/chartsheets/sheet1.xml'),
    ('application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml', '/xl/worksheets/sheet2.xml',),
    ('application/vnd.openxmlformats-officedocument.theme+xml', '/xl/theme/theme1.xml'),
    ('application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml', '/xl/styles.xml'),
    ('application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml', '/xl/sharedStrings.xml'),
    ('application/vnd.openxmlformats-officedocument.drawing+xml', '/xl/drawings/drawing1.xml'),
    ('application/vnd.openxmlformats-officedocument.drawingml.chart+xml','/xl/charts/chart1.xml'),
    ('application/vnd.openxmlformats-officedocument.drawing+xml', '/xl/drawings/drawing2.xml'),
    ('application/vnd.openxmlformats-officedocument.drawingml.chart+xml', '/xl/charts/chart2.xml'),
    ('application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml', '/xl/calcChain.xml'),
    ('application/vnd.openxmlformats-package.core-properties+xml', '/docProps/core.xml'),
    ('application/vnd.openxmlformats-officedocument.extended-properties+xml', '/docProps/app.xml')
    ]


def test_missing_content_type(datadir, DummyArchive):
    from .. workbook import detect_worksheets

    archive = DummyArchive
    datadir.chdir()
    with open("bug181_content_types.xml") as src:
        archive.writestr(ARC_CONTENT_TYPES, src.read())
    with open("bug181_workbook.xml") as src:
        archive.writestr(ARC_WORKBOOK, src.read())
    with open("bug181_workbook.xml.rels") as src:
        archive.writestr(ARC_WORKBOOK_RELS, src.read())
    sheets = list(detect_worksheets(archive))
    assert sheets == [{'path': 'xl/worksheets/sheet1.xml', 'title': 'Sheet 1', 'type':'%s/worksheet' % REL_NS}]


def test_strings_content_type(datadir, DummyArchive):
    from ..workbook import detect_strings

    archive = DummyArchive
    datadir.chdir()
    with open("sharedStrings2.xml") as src:
        archive.writestr(ARC_CONTENT_TYPES, src.read())
    path = detect_strings(archive)
    assert path == 'xl/sharedStrings2.xml'
