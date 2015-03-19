"""
Enclosing chart object. The various chart types are actually child objects.
Will probably need to call this indirectly
"""

from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.descriptors import (
    Bool,
    Float,
    Typed,
    MinMax,
    Integer,
    NoneSet,
)
from openpyxl.descriptors.excel import (
    Percentage,
    ExtensionList
    )

from .text import Tx, TextBody
from .layout import Layout
from .shapes import ShapeProperties
from .legend import Legend
from .marker import PictureOptions, Marker
from .label import DLbl


class Title(Serialisable):

    tx = Typed(expected_type=Tx, allow_none=True)
    layout = Typed(expected_type=Layout, allow_none=True)
    overlay = Bool(nested=True, allow_none=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    txPr = Typed(expected_type=TextBody, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('tx', 'layout', 'overlay', 'spPr', 'txPr', 'extLst')

    def __init__(self,
                 tx=None,
                 layout=None,
                 overlay=None,
                 spPr=None,
                 txPr=None,
                 extLst=None,
                ):
        self.tx = tx
        self.layout = layout
        self.overlay = overlay
        self.spPr = spPr
        self.txPr = txPr


class Surface(Serialisable):

    thickness = Percentage(allow_none=True, nested=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    pictureOptions = Typed(expected_type=PictureOptions, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('thickness', 'spPr', 'pictureOptions', 'extLst')

    def __init__(self,
                 thickness=None,
                 spPr=None,
                 pictureOptions=None,
                 extLst=None,
                ):
        self.thickness = thickness
        self.spPr = spPr
        self.pictureOptions = pictureOptions


class View3D(Serialisable):

    rotX = Integer(allow_none=True, nested=True)
    hPercent = Percentage(allow_none=True, nested=True)
    rotY = Integer(allow_none=True, nested=True)
    depthPercent = Percentage(allow_none=True, nested=True)
    rAngAx = Bool(nested=True, allow_none=True)
    perspective = Integer(nested=True, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('rotX', 'hPercent', 'rotY', 'depthPercent', 'rAngAx', 'perspective', 'extLst')

    def __init__(self,
                 rotX=None,
                 hPercent=None,
                 rotY=None,
                 depthPercent=None,
                 rAngAx=None,
                 perspective=None,
                 extLst=None,
                ):
        self.rotX = rotX
        self.hPercent = hPercent
        self.rotY = rotY
        self.depthPercent = depthPercent
        self.rAngAx = rAngAx
        self.perspective = perspective
        self.extLst = extLst


class PivotFmt(Serialisable):

    idx = Integer(nested=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    txPr = Typed(expected_type=TextBody, allow_none=True)
    marker = Typed(expected_type=Marker, allow_none=True)
    dLbl = Typed(expected_type=DLbl, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('idx', 'spPr', 'txPr', 'marker', 'dLbl', 'extLst')

    def __init__(self,
                 idx=None,
                 spPr=None,
                 txPr=None,
                 marker=None,
                 dLbl=None,
                 extLst=None,
                ):
        self.idx = idx
        self.spPr = spPr
        self.txPr = txPr
        self.marker = marker
        self.dLbl = dLbl


class PivotFmts(Serialisable):

    pivotFmt = Typed(expected_type=PivotFmt, allow_none=True)

    __elements__ = ('pivotFmt',)

    def __init__(self,
                 pivotFmt=None,
                ):
        self.pivotFmt = pivotFmt


class DTable(Serialisable):

    showHorzBorder = Bool(nested=True, allow_none=True)
    showVertBorder = Bool(nested=True, allow_none=True)
    showOutline = Bool(nested=True, allow_none=True)
    showKeys = Bool(nested=True, allow_none=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    txPr = Typed(expected_type=TextBody, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('showHorzBorder', 'showVertBorder', 'showOutline', 'showKeys', 'spPr', 'txPr', 'extLst')

    def __init__(self,
                 showHorzBorder=None,
                 showVertBorder=None,
                 showOutline=None,
                 showKeys=None,
                 spPr=None,
                 txPr=None,
                 extLst=None,
                ):
        self.showHorzBorder = showHorzBorder
        self.showVertBorder = showVertBorder
        self.showOutline = showOutline
        self.showKeys = showKeys
        self.spPr = spPr
        self.txPr = txPr


class PlotArea(Serialisable):

    layout = Typed(expected_type=Layout, allow_none=True)
    dTable = Typed(expected_type=DTable, allow_none=True)
    spPr = Typed(expected_type=ShapeProperties, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('layout', 'dTable', 'spPr', 'extLst')

    def __init__(self,
                 layout=None,
                 dTable=None,
                 spPr=None,
                 extLst=None,
                ):
        self.layout = layout
        self.dTable = dTable
        self.spPr = spPr


class Chart(Serialisable):

    title = Typed(expected_type=Title, allow_none=True)
    autoTitleDeleted = Bool(nested=True, allow_none=True)
    pivotFmts = Typed(expected_type=PivotFmts, allow_none=True)
    view3D = Typed(expected_type=View3D, allow_none=True)
    floor = Typed(expected_type=Surface, allow_none=True)
    sideWall = Typed(expected_type=Surface, allow_none=True)
    backWall = Typed(expected_type=Surface, allow_none=True)
    plotArea = Typed(expected_type=PlotArea, )
    legend = Typed(expected_type=Legend, allow_none=True)
    plotVisOnly = Bool(nested=True, allow_none=True)
    dispBlanksAs = NoneSet(values=(['span', 'gap', 'zero']), nested=True)
    showDLblsOverMax = Bool(nested=True, allow_none=True)
    extLst = Typed(expected_type=ExtensionList, allow_none=True)

    __elements__ = ('title', 'autoTitleDeleted', 'pivotFmts', 'view3D', 'floor', 'sideWall', 'backWall', 'plotArea', 'legend', 'plotVisOnly', 'dispBlanksAs', 'showDLblsOverMax', 'extLst')

    def __init__(self,
                 title=None,
                 autoTitleDeleted=None,
                 pivotFmts=None,
                 view3D=None,
                 floor=None,
                 sideWall=None,
                 backWall=None,
                 plotArea=None,
                 legend=None,
                 plotVisOnly=None,
                 dispBlanksAs=None,
                 showDLblsOverMax=None,
                 extLst=None,
                ):
        self.title = title
        self.autoTitleDeleted = autoTitleDeleted
        self.pivotFmts = pivotFmts
        self.view3D = view3D
        self.floor = floor
        self.sideWall = sideWall
        self.backWall = backWall
        self.plotArea = plotArea
        self.legend = legend
        self.plotVisOnly = plotVisOnly
        self.dispBlanksAs = dispBlanksAs
        self.showDLblsOverMax = showDLblsOverMax
