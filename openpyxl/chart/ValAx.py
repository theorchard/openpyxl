from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Float,
    Set,)

from openpyxl.descriptors.excel import ExtensionList

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


class AxisUnit(Serialisable):

    val = Typed(expected_type=Float(), )

    def __init__(self,
                 val=None,
                ):
        self.val = val


class AxisUnit(Serialisable):

    val = Float()

    def __init__(self,
                 val=None,
                ):
        self.val = val


class CrossBetween(Serialisable):

    val = Set(values=(['between', 'midCat']))

    def __init__(self,
                 val=None,
                ):
        self.val = val


class ValAx(Serialisable):

    crossBetween = Typed(expected_type=CrossBetween, allow_none=True)
    majorUnit = Typed(expected_type=AxisUnit, allow_none=True)
    minorUnit = Typed(expected_type=AxisUnit, allow_none=True)
    dispUnits = Typed(expected_type=DispUnits, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('crossBetween', 'majorUnit', 'minorUnit', 'dispUnits', 'extLst')

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
        self.extLst = extLst

