from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    String,
    Integer,
    Bool,
    Set,
    Float,
)
from openpyxl.descriptors.excel import ExtensionList

from .shapes import ShapeProperties
from .text import TextBody, NumFmt, Tx
from .layout import Layout

class TrendlineLbl(Serialisable):

    layout = Typed(expected_type=Layout, allow_none=True)
    tx = Typed(expected_type=Tx, allow_none=True)
    numFmt = Typed(expected_type=NumFmt, allow_none=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    txPr = Typed(expected_type=TextBody, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    def __init__(self,
                 layout=None,
                 tx=None,
                 numFmt=None,
                 spPr=None,
                 txPr=None,
                 extLst=None,
                ):
        self.layout = layout
        self.tx = tx
        self.numFmt = numFmt
        self.spPr = spPr
        self.txPr = txPr
        self.extLst = extLst



class Period(Serialisable):

    val = Integer()

    def __init__(self,
                 val=None,
                ):
        self.val = val


class Order(Serialisable):

    val = Integer()

    def __init__(self,
                 val=None,
                ):
        self.val = val


class TrendlineType(Serialisable):

    val = Set(values=(['exp', 'linear', 'log', 'movingAvg', 'poly', 'power']))

    def __init__(self,
                 val=None,
                ):
        self.val = val


class Trendline(Serialisable):

    name = String(allow_none=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    trendlineType = Typed(expected_type=TrendlineType, )
    order = Typed(expected_type=Order, allow_none=True)
    period = Typed(expected_type=Period, allow_none=True)
    forward = Float(allow_none=True, nested=True)
    backward = Float(allow_none=True, nested=True)
    intercept = Float(allow_none=True, nested=True)
    dispRSqr = Bool(allow_none=True, nested=True)
    dispEq = Bool(nested=True)
    trendlineLbl = Typed(expected_type=TrendlineLbl, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    def __init__(self,
                 name=None,
                 spPr=None,
                 trendlineType=None,
                 order=None,
                 period=None,
                 forward=None,
                 backward=None,
                 intercept=None,
                 dispRSqr=None,
                 dispEq=None,
                 trendlineLbl=None,
                 extLst=None,
                ):
        self.name = name
        self.spPr = spPr
        self.trendlineType = trendlineType
        self.order = order
        self.period = period
        self.forward = forward
        self.backward = backward
        self.intercept = intercept
        self.dispRSqr = dispRSqr
        self.dispEq = dispEq
        self.trendlineLbl = trendlineLbl
        self.extLst = extLst

