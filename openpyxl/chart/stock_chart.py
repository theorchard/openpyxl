from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Integer,
)
from openpyxl.descriptors.excel import ExtensionList

from .shapes import ShapeProperties
from .chartBase import UpDownBars, ChartLines
from .label import DLbls


class StockChart(Serialisable):

    ser = Typed(expected_type=LineSer, )
    dLbls = Typed(expected_type=DLbls, allow_none=True)
    dropLines = Typed(expected_type=ChartLines, allow_none=True)
    hiLowLines = Typed(expected_type=ChartLines, allow_none=True)
    upDownBars = Typed(expected_type=UpDownBars, allow_none=True)
    axId = Integer(nested=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('ser', 'dLbls', 'dropLines', 'hiLowLines', 'upDownBars', 'axId', 'extLst')

    def __init__(self,
                 ser=None,
                 dLbls=None,
                 dropLines=None,
                 hiLowLines=None,
                 upDownBars=None,
                 axId=None,
                 extLst=None,
                ):
        self.ser = ser
        self.dLbls = dLbls
        self.dropLines = dropLines
        self.hiLowLines = hiLowLines
        self.upDownBars = upDownBars
        self.axId = axId
        self.extLst = extLst
