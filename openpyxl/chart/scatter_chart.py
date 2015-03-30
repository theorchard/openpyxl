from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    NoneSet,
    Bool,
    Integer,
)
from openpyxl.descriptors.excel import ExtensionList

from .series import ScatterSer
from .label import DLbls


class ScatterStyle(Serialisable):

    val = NoneSet(values=(['line', 'lineMarker', 'marker', 'smooth', 'smoothMarker']))

    def __init__(self,
                 val=None,
                ):
        self.val = val


class ScatterChart(Serialisable):

    scatterStyle = Typed(expected_type=ScatterStyle, )
    varyColors = Bool(nested=True, allow_none=True)
    ser = Typed(expected_type=ScatterSer, allow_none=True)
    dLbls = Typed(expected_type=DLbls, allow_none=True)
    axId = Integer(nested=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('scatterStyle', 'varyColors', 'ser', 'dLbls', 'axId', 'extLst')

    def __init__(self,
                 scatterStyle=None,
                 varyColors=None,
                 ser=None,
                 dLbls=None,
                 axId=None,
                 extLst=None,
                ):
        self.scatterStyle = scatterStyle
        self.varyColors = varyColors
        self.ser = ser
        self.dLbls = dLbls
        self.axId = axId
        self.extLst = extLst

