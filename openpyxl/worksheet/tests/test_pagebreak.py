from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import pytest
from openpyxl.tests.helper import compare_xml

from openpyxl.xml.functions import tostring


@pytest.fixture
def Break():
    from ..pagebreak import Break
    return Break


@pytest.fixture
def PageBreak():
    from ..pagebreak import PageBreak
    return PageBreak



class TestBreak:

    pass


class TestPageBreak:

    pass
