from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from openpyxl.compat import safe_string

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Alias,
    Typed,
    Bool,
    Integer,
    Set,
)


class StyleId(Serialisable):
    """
    Format aggregation class

    This is a virtual style composed of references to global format objects
    """
    tagname = "xf"

    numFmtId = Integer()
    number_format = Alias("numFtdId")
    fontId = Integer()
    font = Alias("fontId")
    fillId = Integer()
    fill = Alias("fillId")
    borderId = Integer()
    border = Alias("border")
    xfId = Integer()
    alignment = Integer()
    protection = Integer()
    quotePrefix = Bool(allow_none=True)
    pivotButton = Bool(allow_none=True)
    applyAlignment = Bool(allow_none=True)
    applyProtection = Bool(allow_none=True)

    def __init__(self,
                 numFmtId=0,
                 fontId=0,
                 fillId=0,
                 borderId=0,
                 alignment=0,
                 protection=0,
                 xfId=0,
                 quotePrefix=None,
                 pivotButton=None,
                 applyNumberFormat=None,
                 applyFont=None,
                 applyFill=None,
                 applyBorder=None,
                 applyAlignment=None,
                 applyProtection=None,
                 extLst=None,
                 ):
        self.numFmtId = numFmtId
        self.fontId = fontId
        self.fillId = fillId
        self.borderId = borderId
        self.xfId = xfId
        self.quotePrefix = quotePrefix
        self.pivotButton = pivotButton
        self.alignment = alignment
        self.protection = protection

    @property
    def applyAlignment(self):
        return self.alignment != 0 or None

    @property
    def applyProtection(self):
        return self.protection != 0 or None

    def to_tree(self):
        """
        Alignment and protection objects are implemented as child elements.
        This is a completely different API to other format objects. :-/
        """
        attrs = set(self.__attrs__)
        attrs.discard('alignment')
        attrs.discard('protection')
        attrs.add('applyProtection')
        attrs.add('applyAlignment')
        self.__attrs__ = attrs
        return super(StyleId, self).to_tree()
