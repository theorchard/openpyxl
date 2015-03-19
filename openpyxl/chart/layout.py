from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    NoneSet,
    Float,
    Typed,
)

from openpyxl.descriptors.excel import ExtensionList


class ManualLayout(Serialisable):

    layoutTarget = NoneSet(values=(['inner', 'outer']), nested=True)
    xMode = NoneSet(values=(['edge', 'factor']), nested=True)
    yMode = NoneSet(values=(['edge', 'factor']), nested=True)
    wMode = NoneSet(values=(['edge', 'factor']), nested=True)
    hMode = NoneSet(values=(['edge', 'factor']), nested=True)
    x = Float(nested=True, allow_none=True)
    y = Float(nested=True, allow_none=True)
    w = Float(nested=True, allow_none=True)
    h = Float(nested=True, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('layoutTarget', 'xMode', 'yMode', 'wMode', 'hMode', 'x', 'y', 'w', 'h', 'extLst')

    def __init__(self,
                 layoutTarget=None,
                 xMode=None,
                 yMode=None,
                 wMode=None,
                 hMode=None,
                 x=None,
                 y=None,
                 w=None,
                 h=None,
                 extLst=None,
                ):
        self.layoutTarget = layoutTarget
        self.xMode = xMode
        self.yMode = yMode
        self.wMode = wMode
        self.hMode = hMode
        self.x = x
        self.y = y
        self.w = w
        self.h = h


class Layout(Serialisable):

    tagname = "layout"

    manualLayout = Typed(expected_type=ManualLayout, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('manualLayout', 'extLst')

    def __init__(self,
                 manualLayout=None,
                 extLst=None,
                ):
        self.manualLayout = manualLayout
        self.extLst = extLst
