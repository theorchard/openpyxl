from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    NoneSet,
    Integer,
)

from .layout import Layout
from .shapes import ShapeProperties
from .text import TextBody


class LegendEntry(Serialisable):

    idx = Integer()
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('idx', 'extLst')

    def __init__(self,
                 idx=None,
                 extLst=None,
                ):
        self.idx = idx


class Legend(Serialisable):

    legendPos = NoneSet(values=(['b', 'tr', 'l', 'r', 't']), nested=True)
    legendEntry = Typed(expected_type=LegendEntry, allow_none=True)
    layout = Typed(expected_type=Layout, allow_none=True)
    overlay = Bool(nested=True, allow_none=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    txPr = Typed(expected_type=TextBody, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('legendPos', 'legendEntry', 'layout', 'overlay', 'spPr', 'txPr', 'extLst')

    def __init__(self,
                 legendPos=None,
                 legendEntry=None,
                 layout=None,
                 overlay=None,
                 spPr=None,
                 txPr=None,
                 extLst=None,
                ):
        self.legendPos = legendPos
        self.legendEntry = legendEntry
        self.layout = layout
        self.overlay = overlay
        self.spPr = spPr
        self.txPr = txPr
