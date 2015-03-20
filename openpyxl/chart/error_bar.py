from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Bool,
    Typed,
    Float,
    Set,
)
from .chartBase import *
from .shapes import ShapeProperties

from openpyxl.descriptors.excel import ExtensionList


class ErrValType(Serialisable):

    val = Typed(expected_type=Set(values=(['cust', 'fixedVal', 'percentage', 'stdDev', 'stdErr'])))

    def __init__(self,
                 val=None,
                ):
        self.val = val


class ErrBarType(Serialisable):

    val = Typed(expected_type=Set(values=(['both', 'minus', 'plus'])))

    def __init__(self,
                 val=None,
                ):
        self.val = val


class ErrDir(Serialisable):

    val = Typed(expected_type=Set(values=(['x', 'y'])))

    def __init__(self,
                 val=None,
                ):
        self.val = val


class ErrBars(Serialisable):

    errDir = Typed(expected_type=ErrDir, allow_none=True)
    errBarType = Typed(expected_type=ErrBarType, )
    errValType = Typed(expected_type=ErrValType, )
    noEndCap = Bool(nested=True, allow_none=True)
    plus = Typed(expected_type=NumDataSource, allow_none=True)
    minus = Typed(expected_type=NumDataSource, allow_none=True)
    val = Float(allow_none=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    def __init__(self,
                 errDir=None,
                 errBarType=None,
                 errValType=None,
                 noEndCap=None,
                 plus=None,
                 minus=None,
                 val=None,
                 spPr=None,
                 extLst=None,
                ):
        self.errDir = errDir
        self.errBarType = errBarType
        self.errValType = errValType
        self.noEndCap = noEndCap
        self.plus = plus
        self.minus = minus
        self.val = val
        self.spPr = spPr
        self.extLst = extLst
