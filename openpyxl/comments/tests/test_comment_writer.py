from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from openpyxl.compat import zip
from openpyxl.workbook import Workbook
from openpyxl.worksheet import Worksheet
from openpyxl.writer.comments import CommentWriter
from openpyxl.comments import Comment
from openpyxl.tests.helper import compare_xml
from openpyxl.xml.functions import fromstring, tostring
from openpyxl.xml.constants import SHEET_MAIN_NS
from openpyxl.writer.comments import vmlns, excelns

def _create_ws():
    wb = Workbook()
    ws = Worksheet(wb)
    comment1 = Comment("text", "author")
    comment2 = Comment("text2", "author2")
    comment3 = Comment("text3", "author3")
    ws["B2"].comment = comment1
    ws["C7"].comment = comment2
    ws["D9"].comment = comment3
    return ws, comment1, comment2, comment3

def test_comment_writer_init():
    ws, comment1, comment2, comment3 = _create_ws()
    cw = CommentWriter(ws)
    assert set(cw.authors) == set(["author", "author2", "author3"])
    assert set(cw.comments) == set([comment1, comment2, comment3])

def test_write_comments(datadir):
    datadir.chdir()
    ws = _create_ws()[0]
    cw = CommentWriter(ws)
    content = cw.write_comments()
    with open('comments1.xml') as expected:
        correct = fromstring(expected.read())
    check = fromstring(content)
    # check top-level elements have the same name
    for i, j in zip(correct.getchildren(), check.getchildren()):
        assert i.tag == j.tag

    correct_comments = correct.find('{%s}commentList' % SHEET_MAIN_NS).getchildren()
    check_comments = check.find('{%s}commentList' % SHEET_MAIN_NS).getchildren()
    correct_authors = correct.find('{%s}authors' % SHEET_MAIN_NS).getchildren()
    check_authors = check.find('{%s}authors' % SHEET_MAIN_NS).getchildren()

    # replace author ids with author names
    for i in correct_comments:
        i.attrib["authorId"] = correct_authors[int(i.attrib["authorId"])].text
    for i in check_comments:
        i.attrib["authorId"] = check_authors[int(i.attrib["authorId"])].text

    # sort the comment list
    correct_comments.sort(key=lambda tag: tag.attrib["ref"])
    check_comments.sort(key=lambda tag: tag.attrib["ref"])
    correct.find('{%s}commentList' % SHEET_MAIN_NS)[:] = correct_comments
    check.find('{%s}commentList' % SHEET_MAIN_NS)[:] = check_comments

    # sort the author list
    correct_authors.sort(key=lambda tag: tag.text)
    check_authors.sort(key=lambda tag:tag.text)
    correct.find('{%s}authors' % SHEET_MAIN_NS)[:] = correct_authors
    check.find('{%s}authors' % SHEET_MAIN_NS)[:] = check_authors

    diff = compare_xml(tostring(correct), tostring(check))
    assert diff is None, diff

def test_write_comments_vml(datadir):
    datadir.chdir()
    ws = _create_ws()[0]
    cw = CommentWriter(ws)
    content = cw.write_comments_vml()
    with open('commentsDrawing1.vml') as expected:
        correct = fromstring(expected.read())
    check = fromstring(content)
    correct_ids = []
    correct_coords = []
    check_ids = []
    check_coords = []

    for i in correct.findall("{%s}shape" % vmlns):
        correct_ids.append(i.attrib["id"])
        row = i.find("{%s}ClientData" % excelns).find("{%s}Row" % excelns).text
        col = i.find("{%s}ClientData" % excelns).find("{%s}Column" % excelns).text
        correct_coords.append((row,col))
        # blank the data we are checking separately
        i.attrib["id"] = "0"
        i.find("{%s}ClientData" % excelns).find("{%s}Row" % excelns).text="0"
        i.find("{%s}ClientData" % excelns).find("{%s}Column" % excelns).text="0"

    for i in check.findall("{%s}shape" % vmlns):
        check_ids.append(i.attrib["id"])
        row = i.find("{%s}ClientData" % excelns).find("{%s}Row" % excelns).text
        col = i.find("{%s}ClientData" % excelns).find("{%s}Column" % excelns).text
        check_coords.append((row,col))
        # blank the data we are checking separately
        i.attrib["id"] = "0"
        i.find("{%s}ClientData" % excelns).find("{%s}Row" % excelns).text="0"
        i.find("{%s}ClientData" % excelns).find("{%s}Column" % excelns).text="0"

    assert set(correct_coords) == set(check_coords)
    assert set(correct_ids) == set(check_ids)
    diff = compare_xml(tostring(correct), tostring(check))
    assert diff is None, diff


def test_write_only_cell_vml(datadir):
    from openpyxl.xml.functions import Element, tostring
    datadir.chdir()
    wb = Workbook()
    ws = wb.active
    cell = ws['A1'] # write-only cells are always A1
    cell.comment = Comment("Some text", "an author")
    cell.col_idx = 2
    cell.row = 2

    writer = CommentWriter(ws)
    root = Element("root")
    xml = writer._write_comment_shape(cell.comment, 1)
    xml = tostring(xml)
    expected = """
    <v:shape
    xmlns:v="urn:schemas-microsoft-com:vml"
    xmlns:x="urn:schemas-microsoft-com:office:excel"
    xmlns:o="urn:schemas-microsoft-com:office:office"
    fillcolor="#ffffe1"
    id="_x0000_s0001"
    style="position:absolute; margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;z-index:1;visibility:hidden"
    type="#_x0000_t202"
    o:insetmode="auto">
      <v:fill color2="#ffffe1"/>
      <v:shadow color="black" obscured="t"/>
      <v:path o:connecttype="none"/>
      <v:textbox style="mso-direction-alt:auto">
        <div style="text-align:left"/>
      </v:textbox>
      <x:ClientData ObjectType="Note">
        <x:MoveWithCells/>
        <x:SizeWithCells/>
        <x:AutoFill>False</x:AutoFill>
        <x:Row>1</x:Row>
        <x:Column>1</x:Column>
      </x:ClientData>
    </v:shape>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff
