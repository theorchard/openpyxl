from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl


from zipfile import ZipFile

from openpyxl.workbook import Workbook
from openpyxl.worksheet import Worksheet
from openpyxl.reader import comments
from openpyxl.reader.excel import load_workbook
from openpyxl.xml.functions import fromstring

import pytest


def test_get_author_list():
    xml = """<?xml version="1.0" standalone="yes"?><comments
    xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors>
    <author>Cuke</author><author>Not Cuke</author></authors><commentList>
    </commentList></comments>"""
    assert comments._get_author_list(fromstring(xml)) == ['Cuke', 'Not Cuke']


def test_read_comments():
    xml = """<?xml version="1.0" standalone="yes"?>
    <comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><authors>
    <author>Cuke</author><author>Not Cuke</author></authors><commentList><comment ref="A1"
    authorId="0" shapeId="0"><text><r><rPr><b/><sz val="9"/><color indexed="81"/><rFont
    val="Tahoma"/><charset val="1"/></rPr><t>Cuke:\n</t></r><r><rPr><sz val="9"/><color
    indexed="81"/><rFont val="Tahoma"/><charset val="1"/></rPr>
    <t xml:space="preserve">First Comment</t></r></text></comment><comment ref="D1" authorId="0" shapeId="0">
    <text><r><rPr><b/><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><charset val="1"/>
    </rPr><t>Cuke:\n</t></r><r><rPr><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/>
    <charset val="1"/></rPr><t xml:space="preserve">Second Comment</t></r></text></comment>
    <comment ref="A2" authorId="1" shapeId="0"><text><r><rPr><b/><sz val="9"/><color
    indexed="81"/><rFont val="Tahoma"/><charset val="1"/></rPr><t>Not Cuke:\n</t></r><r><rPr>
    <sz val="9"/><color indexed="81"/><rFont val="Tahoma"/><charset val="1"/></rPr>
    <t xml:space="preserve">Third Comment</t></r></text></comment></commentList></comments>"""
    wb = Workbook()
    ws = Worksheet(wb)
    comments.read_comments(ws, xml)
    comments_expected = [['A1', 'Cuke', 'Cuke:\nFirst Comment'],
                         ['D1', 'Cuke', 'Cuke:\nSecond Comment'],
                         ['A2', 'Not Cuke', 'Not Cuke:\nThird Comment']
                        ]
    for cell, author, text in comments_expected:
        assert ws.cell(coordinate=cell).comment.author == author
        assert ws.cell(coordinate=cell).comment.text == text
        assert ws.cell(coordinate=cell).comment._parent == ws.cell(coordinate=cell)


def test_get_comments_file(datadir):
    datadir.chdir()
    archive = ZipFile('comments.xlsx')
    valid_files = archive.namelist()
    assert comments.get_comments_file('sheet1.xml', archive, valid_files) == 'xl/comments1.xml'
    assert comments.get_comments_file('sheet3.xml', archive, valid_files) == 'xl/comments2.xml'
    assert comments.get_comments_file('sheet2.xml', archive, valid_files) is None


def test_comments_cell_association(datadir):
    datadir.chdir()
    wb = load_workbook('comments.xlsx')
    assert wb['Sheet1'].cell(coordinate="A1").comment.author == "Cuke"
    assert wb['Sheet1'].cell(coordinate="A1").comment.text == "Cuke:\nFirst Comment"
    assert wb['Sheet2'].cell(coordinate="A1").comment is None
    assert wb['Sheet1'].cell(coordinate="D1").comment.text == "Cuke:\nSecond Comment"


@pytest.mark.xfail
def test_comments_with_iterators(datadir):
    datadir.chdir()
    wb = load_workbook('comments.xlsx', use_iterators=True)
    ws = wb['Sheet1']
    assert ws.cell(coordinate="A1").comment.author == "Cuke"
    assert ws.cell(coordinate="A1").comment.text == "Cuke:\nFirst Comment"
