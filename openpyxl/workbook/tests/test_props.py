from __future__ import absolute_import
# copyright openpyxl 2014

import datetime

import pytest
from lxml.etree import fromstring
from openpyxl.tests.helper import compare_xml


def test_ctor():
    from .. properties import DocumentProperties
    dt = datetime.datetime(2014, 10, 12, 10, 35, 36)
    props = DocumentProperties(created=dt, modified=dt)
    assert dict(props) == {'created': '2014-10-12 10:35:36', 'modified':
                           '2014-10-12 10:35:36', 'creator': 'openpyxl',}


@pytest.fixture()
def SampleProperties():
    from .. properties import DocumentProperties
    props = DocumentProperties()
    props.keywords = "one, two, three"
    props.created = datetime.datetime(2010, 4, 1, 20, 30, 00)
    props.modified = datetime.datetime(2010, 4, 5, 14, 5, 30)
    props.lastPrinted = datetime.datetime(2014, 10, 14, 10, 30)
    props.category = "The category"
    props.contentStatus = "The status"
    props.creator = 'TEST_USER'
    props.lastModifiedBy = "SOMEBODY"
    props.revision = "0"
    props.version = "2.5"
    props.description = "The description"
    props.identifier = "The identifier"
    props.language = "The language"
    props.subject = "The subject"
    props.title = "The title"
    return props


def test_dict_interface(SampleProperties):
    assert dict(SampleProperties) == {
        'created': '2010-04-01 20:30:00',
        'creator': 'TEST_USER',
        'lastModifiedBy': 'SOMEBODY',
        'modified':'2010-04-05 14:05:30',
        'category': 'The category',
        'contentStatus': 'The status',
        'description': 'The description',
        'identifier': 'The identifier',
        'language': 'The language',
        'lastPrinted': '2014-10-14 10:30:00',
        'revision': '0',
        'subject': 'The subject',
        'title': 'The title',
        'version': '2.5',
        'keywords': 'one, two, three',
                           }


def test_write_properties_core(datadir, SampleProperties):
    from .. properties import write_properties
    datadir.chdir()

    content = write_properties(SampleProperties)
    with open('core.xml') as expected:
        diff = compare_xml(content, expected.read())
    assert diff is None, diff


def test_read_properties_core(datadir, SampleProperties):
    from .. properties import read_properties
    datadir.chdir()

    with open("core.xml") as src:
        content = src.read()
    props = read_properties(content)
    assert dict(props) == dict(SampleProperties)
