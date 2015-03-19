from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Float,
    NoneSet,
    Bool,
    Integer,
    MinMax,
    NoneSet,
)

from openpyxl.descriptors.excel import ExtensionList, Percentage

from .layout import Layout
from .text import Tx, TextBody
from .shapes import ShapeProperties


class DispUnitsLbl(Serialisable):

    layout = Typed(expected_type=Layout, allow_none=True)
    tx = Typed(expected_type=Tx, allow_none=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    txPr = Typed(expected_type=TextBody, allow_none=True)

    __elements__ = ('layout', 'tx', 'spPr', 'txPr')

    def __init__(self,
                 layout=None,
                 tx=None,
                 spPr=None,
                 txPr=None,
                ):
        self.layout = layout
        self.tx = tx
        self.spPr = spPr
        self.txPr = txPr


class DispUnits(Serialisable):

    dispUnitsLbl = Typed(expected_type=DispUnitsLbl, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('dispUnitsLbl', 'extLst')

    def __init__(self,
                 dispUnitsLbl=None,
                 extLst=None,
                ):
        self.dispUnitsLbl = dispUnitsLbl
        self.extLst = extLst


class ValAx(Serialisable):

    tagname = "valAx"

    crossBetween = NoneSet(values=(['between', 'midCat']), nested=True)
    majorUnit = Float(allow_none=True, nested=True)
    minorUnit = Float(allow_none=True, nested=True)
    dispUnits = Typed(expected_type=DispUnits, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ()
    __nested__ = ('crossBetween', 'majorUnit', 'minorUnit', 'dispUnits',)

    def __init__(self,
                 crossBetween=None,
                 majorUnit=None,
                 minorUnit=None,
                 dispUnits=None,
                 extLst=None,
                ):
        self.crossBetween = crossBetween
        self.majorUnit = majorUnit
        self.minorUnit = minorUnit
        self.dispUnits = dispUnits


class CatAx(Serialisable):

    tagname = "catAx"

    auto = Bool(nested=True, allow_none=True)
    lblAlgn = NoneSet(values=(['ctr', 'l', 'r']), nested=True)
    lblOffset = MinMax(min=0, max=1000, nested=True,)
    tickLblSkip = Integer(allow_none=True, nested=True)
    tickMarkSkip = Integer(allow_none=True, nested=True)
    noMultiLvlLbl = Bool(nested=True, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ()
    __nested__ = ('auto', 'lblAlgn', 'lblOffset', 'tickLblSkip', 'tickMarkSkip', 'noMultiLvlLbl')

    def __init__(self,
                 auto=None,
                 lblAlgn=None,
                 lblOffset=100,
                 tickLblSkip=None,
                 tickMarkSkip=None,
                 noMultiLvlLbl=None,
                 extLst=None,
                ):
        self.auto = auto
        self.lblAlgn = lblAlgn
        self.lblOffset = lblOffset
        self.tickLblSkip = tickLblSkip
        self.tickMarkSkip = tickMarkSkip
        self.noMultiLvlLbl = noMultiLvlLbl


class DateAx(Serialisable):

    auto = Bool(nested=True, allow_none=True)
    lblOffset = Percentage(allow_none=True, nested=True)
    baseTimeUnit = NoneSet(values=(['days', 'months', 'years']), nested=True)
    majorUnit = Float(allow_none=True, nested=True)
    majorTimeUnit = NoneSet(values=(['days', 'months', 'years']), nested=True)
    minorUnit = Float(allow_none=True, nested=True)
    minorTimeUnit = NoneSet(values=(['days', 'months', 'years']), nested=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('auto', 'lblOffset', 'baseTimeUnit', 'majorUnit', 'majorTimeUnit', 'minorUnit', 'minorTimeUnit', 'extLst')

    def __init__(self,
                 auto=None,
                 lblOffset=None,
                 baseTimeUnit=None,
                 majorUnit=None,
                 majorTimeUnit=None,
                 minorUnit=None,
                 minorTimeUnit=None,
                 extLst=None,
                ):
        self.auto = auto
        self.lblOffset = lblOffset
        self.baseTimeUnit = baseTimeUnit
        self.majorUnit = majorUnit
        self.majorTimeUnit = majorTimeUnit
        self.minorUnit = minorUnit
        self.minorTimeUnit = minorTimeUnit


class SerAx(Serialisable):

    tickLblSkip = Integer(allow_none=True, nested=True)
    tickMarkSkip = Integer(allow_none=True, nested=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('tickLblSkip', 'tickMarkSkip', 'extLst')

    def __init__(self,
                 tickLblSkip=None,
                 tickMarkSkip=None,
                 extLst=None,
                ):
        self.tickLblSkip = tickLblSkip
        self.tickMarkSkip = tickMarkSkip
