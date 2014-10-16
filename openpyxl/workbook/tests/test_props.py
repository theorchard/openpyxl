from __future__ import absolute_import
# copyright openpyxl 2014

import datetime

import pytest
from lxml.etree import fromstring
from openpyxl.tests.helper import compare_xml
from openpyxl.tests.schema import core_props_schema


@pytest.fixture()
def DocumentProperties():
    from .. properties import DocumentProperties
    return DocumentProperties


def test_ctor(DocumentProperties):
    dt = datetime.datetime(2014, 10, 12, 10, 35, 36)
    props = DocumentProperties(created=dt, modified=dt)
    assert dict(props) == {'created': '2014-10-12 10:35:36', 'modified':
                           '2014-10-12 10:35:36', 'creator': 'openpyxl',}


def test_write_properties_core(datadir, DocumentProperties):
    from .. properties import write_properties
    datadir.chdir()
    prop = DocumentProperties()
    prop.creator = 'TEST_USER'
    prop.last_modified_by = 'SOMEBODY'
    prop.created = datetime.datetime(2010, 4, 1, 20, 30, 00)
    prop.modified = datetime.datetime(2010, 4, 5, 14, 5, 30)
    prop.lastPrinted = datetime.datetime(2014, 10, 14, 10, 30)
    content = write_properties(prop)
    with open('core.xml') as expected:
        diff = compare_xml(content, expected.read())
    assert diff is None, diff


def test_validate_schema(DocumentProperties):
    props = DocumentProperties()
    props.keywords = "one, two, three"
    props.created = datetime.datetime(2010, 4, 1, 20, 30, 00)
    props.modified = datetime.datetime(2010, 4, 5, 14, 5, 30)
    props.lastPrinted = datetime.datetime(2014, 10, 14, 10, 30)
    props.category = "The category"
    props.contentStatus = "The status"
    prop.creator = 'TEST_USER'
    props.lastModifiedBy = "SOMEBODY"
    props.revision = "0"
    props.version = "2.5"
    props.description = "The description"
    props.identifier = "The identifier"
    props.language = "The language"
    props.subject = "The subject"
    props.title = "The title"
    from .. properties import write_properties
    xml = write_properties(props)
    root = fromstring(xml)
    core_props_schema.assertValid(root)


def test_read_properties_core(datadir):
    from .. properties import read_properties
    datadir.chdir()
    with open("sample_core_properties.xml") as src:
        content = src.read()
    prop = read_properties(content)
    assert prop.creator == '*.*'
    assert prop.last_modified_by == 'Charlie Clark'
    assert prop.created == datetime.datetime(2010, 4, 9, 20, 43, 12)
    assert prop.modified ==  datetime.datetime(2014, 1, 2, 14, 53, 6)


def test_read_properties_libreeoffice(datadir):
    from .. properties import read_properties
    datadir.chdir()
    with open("libre_office_properties.xml") as src:
        content = src.read()
    prop = read_properties(content)
    assert prop.revision == "0"
