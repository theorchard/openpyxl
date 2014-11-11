from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl


from openpyxl.descriptors import Strict
from openpyxl.xml.constants import SHEET_MAIN_NS


class NamedStyle(object):

    tag = '{%s}cellXfs' % SHEET_MAIN_NS

    """
    Named and editable styles
    """

    def __init__(self, name):
        self.name = name
