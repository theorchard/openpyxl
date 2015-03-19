from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Float,
    NoneSet,
    Bool,
    Integer,
    MinMax,
    NoneSet,
    Set,
)

from openpyxl.descriptors.excel import ExtensionList, Percentage

from openpyxl.styles.differential import NumFmt

from .layout import Layout
from .text import Tx, TextBody
from .shapes import ShapeProperties
from .chartBase import ChartLines
from ._chart import Title


class Scaling(Serialisable):

    tagname = "scaling"

    logBase = Float(allow_none=True, nested=True)
    orientation = Set(values=(['maxMin', 'minMax']), nested=True)
    max = Float(nested=True, allow_none=True)
    min = Float(nested=True, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ()
    __nested__ = ('logBase', 'orientation', 'max', 'min',)

    def __init__(self,
                 logBase=None,
                 orientation="minMax",
                 max=None,
                 min=None,
                 extLst=None,
                ):
        self.logBase = logBase
        self.orientation = orientation
        self.max = max
        self.min = min


class _BaseAxis(Serialisable):

    axId = Integer(nested=True)
    scaling = Typed(expected_type=Scaling)
    delete = Bool(nested=True, allow_none=True)
    axPos = NoneSet(values=(['b', 'l', 'r', 't']))
    majorGridlines = Typed(expected_type=ChartLines, allow_none=True)
    minorGridlines = Typed(expected_type=ChartLines, allow_none=True)
    title = Typed(expected_type=Title, allow_none=True)
    numFmt = Typed(expected_type=NumFmt, allow_none=True)
    majorTickMark = NoneSet(values=(['cross', 'in', 'out']), nested=True)
    minorTickMark = NoneSet(values=(['cross', 'in', 'out']), nested=True)
    tickLblPos = NoneSet(values=(['high', 'low', 'nextTo']), nested=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    txP = Typed(expected_type=TextBody, allow_none=True)
    crossAx = Integer(nested=True) # references other axis
    crosses = NoneSet(values=(['autoZero', 'max', 'min']), nested=True)
    crossesAt = Float(nested=True, allow_none=True)

    # crosses & crossesAt are mutually exclusive

    __nested__ = ('axId', 'delete', 'majorTickMark', 'minorTickMark',
                  'tickLblPos', 'crossAx', 'crosses', 'crossesAt')
    __elements__ = ('majorGridlines', 'minorGridlines', 'numFmt', 'scaling',
                    'spPr', 'title', 'txP')

    def __init__(self,
                 axId=None,
                 scaling=None,
                 delete=None,
                 axPos=None,
                 majorGridlines=None,
                 minorGridlines=None,
                 title=None,
                 numFmt=None,
                 majorTickMark=None,
                 minorTickMark=None,
                 tickLblPos=None,
                 spPr=None,
                 txP= None,
                 crossAx=None,
                 crosses=None,
                 crossesAt=None,
                ):
        self.axId = axId
        if scaling is None:
            self.scaling = Scaling()
        self.delete = delete
        self.axPos = axPos
        self.majorGridlines = majorGridlines
        self.minorGridlines = minorGridlines
        self.title = title
        self.numFmt = numFmt
        self.majorTickMark = majorTickMark
        self.minorTickMark = minorTickMark
        self.tickLblPos = tickLblPos
        self.spPr = spPr
        self.txP = txP
        self.crossAx = crossAx
        self.crosses = crosses
        self.crossesAt = None


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
