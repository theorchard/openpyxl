from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Typed,
    Set,
    Bool,
    Integer,
)
from openpyxl.descriptors.excel import ExtensionList

from .series import RadarSer
from .label import DLbls


class RadarStyle(Serialisable):

    val = Set(values=(['standard', 'marker', 'filled']))

    def __init__(self,
                 val=None,
                ):
        self.val = val


class RadarChart(Serialisable):

    radarStyle = Typed(expected_type=RadarStyle, )
    varyColors = Bool(nested=True, allow_none=True)
    ser = Typed(expected_type=RadarSer, allow_none=True)
    dLbls = Typed(expected_type=DLbls, allow_none=True)
    axId = Integer(nested=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('radarStyle', 'varyColors', 'ser', 'dLbls', 'axId', 'extLst')

    def __init__(self,
                 radarStyle=None,
                 varyColors=None,
                 ser=None,
                 dLbls=None,
                 axId=None,
                 extLst=None,
                ):
        self.radarStyle = radarStyle
        self.varyColors = varyColors
        self.ser = ser
        self.dLbls = dLbls
        self.axId = axId
        self.extLst = extLst

