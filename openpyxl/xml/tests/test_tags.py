# Copyright (c) 2010-2014 openpyxl

from io import BytesIO
import pytest

from openpyxl.xml.functions import start_tag, end_tag, tag, XMLGenerator

@pytest.fixture
def doc():
    return BytesIO()

@pytest.fixture
def root(doc):
    return XMLGenerator(doc)


class TestSimpleTag:

    def test_start_tag(self, doc, root):
        start_tag(root, "start")
        assert doc.getvalue() == b"<start>"

    def test_end_tag(self, doc, root):
        """"""
        end_tag(root, "blah")
        assert doc.getvalue() == b"</blah>"


class TestTagBody:

    def test_start_tag(self, doc, root):
        start_tag(root, "start", body="just words")
        assert doc.getvalue() == b"<start>just words"

    def test_end_tag(self, doc, root):
        end_tag(root, "end")
        assert doc.getvalue() == b"</end>"


def test_start_tag_attrs(doc, root):
    start_tag(root, "start", {'width':"10"})
    assert doc.getvalue() == b"""<start width="10">"""


def test_tag(doc, root):
    tag(root, "start", {'height':"10"}, "words")
    assert doc.getvalue() == b"""<start height="10">words</start>"""
