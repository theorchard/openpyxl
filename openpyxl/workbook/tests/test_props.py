from __future__ import absolute_import
# copyright openpyxl 2014

import datetime

import pytest


@pytest.fixture()
def DocumentProperties():
    from .. properties import DocumentProperties
    return DocumentProperties


def test_ctor(DocumentProperties):
    dt = datetime.datetime(2014, 10, 12, 10, 35, 36)
    props = DocumentProperties(created=dt, modified=dt)
    assert dict(props) == {'created': '2014-10-12 10:35:36', 'modified': '2014-10-12 10:35:36', 'creator': 'openpyxl'}
