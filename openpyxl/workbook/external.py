from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

from openpyxl.collections import IndexedList
from openpyxl.descriptors import String, Strict
from openpyxl.xml.constants import SHEET_MAIN_NS
from openpyxl.xml.functions import fromstring, safe_iterator

"""Manage links to external Workbooks"""


class ExternalRelationship(object):

    """
    Map the relationship of one workbook to another
    """

    def __init__(self, Target, TargetMode):
        self.Target = Target
        self.TargetMode = TargetMode


class ExternalRange(Strict):

    """
    Map external named ranges
    NB. the specification for these is different to named ranges within a workbook
    See 18.14.5
    """

    name = String()
    refersTo = String()
    sheetId = String(allow_none=True)

    def __init__(self, name, refersTo, sheetId=None):
        self.name = name
        self.refersTo = refersTo
        self.sheetId = sheetId


def parse_names(xml):
    tree = fromstring(xml)
    book = tree.find('{%s}externalBook' % SHEET_MAIN_NS)
    names = book.find('{%s}definedNames' % SHEET_MAIN_NS)
    for __n in safe_iterator(names, '{%s}definedName' % SHEET_MAIN_NS):
        yield ExternalRange(**dict(__n.attrib))
