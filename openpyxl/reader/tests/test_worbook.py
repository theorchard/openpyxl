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


@pytest.mark.parametrize("excel_file, expected", [
    ("bug137.xlsx", [
        {'path': 'xl/worksheets/sheet1.xml', 'title': 'Sheet1'}
        ]
     ),
    ("contains_chartsheets.xlsx", [
        {'path': 'xl/worksheets/sheet1.xml', 'title': 'data'},
        {'path': 'xl/worksheets/sheet2.xml', 'title': 'moredata'},
        ]),
    ("bug304.xlsx", [
    {'path': 'xl/worksheets/sheet3.xml', 'title': 'Sheet1'},
    {'path': 'xl/worksheets/sheet2.xml', 'title': 'Sheet2'},
    {'path': 'xl/worksheets/sheet.xml', 'title': 'Sheet3'},
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
        "rId1": {'path': 'xl/chartsheets/sheet1.xml'},
        "rId2": {'path': 'xl/worksheets/sheet1.xml'},
        "rId3": {'path': 'xl/theme/theme1.xml'},
        "rId4": {'path': 'xl/styles.xml'},
        "rId5": {'path': 'xl/sharedStrings.xml'}
    }),
    ("bug304.xlsx", {
        'rId1': {'path': 'xl/worksheets/sheet3.xml'},
        'rId2': {'path': 'xl/worksheets/sheet2.xml'},
        'rId3': {'path': 'xl/worksheets/sheet.xml'},
        'rId4': {'path': 'xl/theme/theme.xml'},
        'rId5': {'path': 'xl/styles.xml'},
        'rId6': {'path': '../customXml/item1.xml'},
        'rId7': {'path': '../customXml/item2.xml'},
        'rId8': {'path': '../customXml/item3.xml'}
    }),
]
                         )
def test_read_rels(datadir, excel_file, expected):
    from openpyxl.reader.workbook import read_rels

    datadir.chdir()
    archive = ZipFile(excel_file)
    assert dict(read_rels(archive)) == expected


@pytest.mark.parametrize("excel_file, expected", [
    ("bug137.xlsx",
     [
         ("rId1", "Chart1", None),
         ("rId2", "Sheet1", None),
     ]
     ),
    ("bug304.xlsx",
     [
         ('rId1', 'Sheet1', None),
         ('rId2', 'Sheet2', None),
         ('rId3', 'Sheet3', None),
     ]
     )
])
def test_read_sheets(datadir, excel_file, expected):
    from openpyxl.reader.workbook import read_sheets

    datadir.chdir()
    archive = ZipFile(excel_file)
    assert list(read_sheets(archive)) == expected


def test_read_content_types(datadir):
    from openpyxl.reader.workbook import read_content_types

    datadir.chdir()
    archive = ZipFile("contains_chartsheets.xlsx")
    assert list(read_content_types(archive)) == [
    ('/xl/workbook.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'),
    ('/xl/worksheets/sheet1.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'),
    ('/xl/chartsheets/sheet1.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml'),
    ('/xl/worksheets/sheet2.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'),
    ('/xl/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml'),
    ('/xl/styles.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'),
    ('/xl/sharedStrings.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'),
    ('/xl/drawings/drawing1.xml', 'application/vnd.openxmlformats-officedocument.drawing+xml'),
    ('/xl/charts/chart1.xml', 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'),
    ('/xl/drawings/drawing2.xml', 'application/vnd.openxmlformats-officedocument.drawing+xml'),
    ('/xl/charts/chart2.xml', 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'),
    ('/xl/calcChain.xml', 'application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml'),
    ('/docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml'),
    ('/docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml')
    ]
