from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from openpyxl.comments import Comment
from openpyxl.workbook import Workbook
from openpyxl.worksheet import Worksheet

def test_init():
    wb = Workbook()
    ws = Worksheet(wb)
    c = Comment("text", "author")
    ws.cell(coordinate="A1").comment = c
    assert c._parent == ws.cell(coordinate="A1")
    assert c.text == "text"
    assert c.author == "author"
