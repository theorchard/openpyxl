from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

from openpyxl.collections import IndexedList
from openpyxl.descriptors import String

"""Manage links to external Workbooks"""


class ExternalRelationship(object):

    """
    Map the relationship of one workbook to another
    """

    def __init__(self, Target, TargetMode):
        self.Target = Target
        self.TargetMode = TargetMode


class ExternalLink(object):

    """
    Map the links to named ranges in an external workbook
    """

    def __init__(self):
        self.names = IndexedList()


class ExternalRange(object):

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
