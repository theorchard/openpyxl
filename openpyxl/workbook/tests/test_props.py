from __future__ import absolute_import
# copyright openpyxl 2014

import datetime

import pytest
from openpyxl.tests.helper import compare_xml


@pytest.fixture()
def DocumentProperties():
    from .. properties import DocumentProperties
    return DocumentProperties


def test_ctor(DocumentProperties):
    dt = datetime.datetime(2014, 10, 12, 10, 35, 36)
    props = DocumentProperties(created=dt, modified=dt)
    assert dict(props) == {'created': '2014-10-12 10:35:36', 'modified': '2014-10-12 10:35:36', 'creator': 'openpyxl'}


def test_write_properties_core(datadir, DocumentProperties):
    from .. properties import write_properties
    datadir.chdir()
    prop = DocumentProperties()
    prop.creator = 'TEST_USER'
    prop.last_modified_by = 'SOMEBODY'
    prop.created = datetime.datetime(2010, 4, 1, 20, 30, 00)
    prop.modified = datetime.datetime(2010, 4, 5, 14, 5, 30)
    content = write_properties(prop)
    with open('core.xml') as expected:
        diff = compare_xml(content, expected.read())
    assert diff is None, diff

