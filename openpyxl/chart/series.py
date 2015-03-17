from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    String,
    Integer,
)
from openpyxl.descriptors.excel import ExtensionList

from .shapes import ShapeProperties

class StrVal(Serialisable):

    idx = Integer()
    v = Typed(expected_type=String(), )

    def __init__(self,
                 idx=None,
                 v=None,
                ):
        self.idx = idx
        self.v = v


class StrData(Serialisable):

    ptCount = Integer(allow_none=True, nested=True)
    pt = Typed(expected_type=StrVal, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('ptCount', 'pt', 'extLst')

    def __init__(self,
                 ptCount=None,
                 pt=None,
                 extLst=None,
                ):
        self.ptCount = ptCount
        self.pt = pt
        self.extLst = extLst


class StrRef(Serialisable):

    f = Typed(expected_type=String, )
    strCache = Typed(expected_type=StrData, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('f', 'strCache', 'extLst')

    def __init__(self,
                 f=None,
                 strCache=None,
                 extLst=None,
                ):
        self.f = f
        self.strCache = strCache
        self.extLst = extLst


class SerTx(Serialisable):

    strRef = Typed(expected_type=StrRef)


class SerShared(Serialisable):

    idx = Integer(nested=True)
    order = Integer(nested=True)
    tx = Typed(expected_type=SerTx, allow_none=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)

    __elements__ = ('idx', 'order', 'tx', 'spPr')

    def __init__(self,
                 idx=None,
                 order=None,
                 tx=None,
                 spPr=None,
                ):
        self.idx = idx
        self.order = order
        self.tx = tx
        self.spPr = spPr

