from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from openpyxl.descriptors import Bool, Integer, String, Set, Float
from openpyxl.descriptors.serialisable import Serialisable

class SheetView(Serialisable):

    """Information about the visible portions of this sheet."""

    tagname = "sheetView"

    windowProtection = Bool(allow_none=True)
    showFormulas = Bool(allow_none=True)
    showGridLines = Bool(allow_none=True)
    showRowColHeaders = Bool(allow_none=True)
    showZeros = Bool(allow_none=True)
    rightToLeft = Bool(allow_none=True)
    tabSelected = Bool(allow_none=True)
    showRuler = Bool(allow_none=True)
    showOutlineSymbols = Bool(allow_none=True)
    defaultGridColor = Bool(allow_none=True)
    showWhiteSpace = Bool(allow_none=True)
    view = Set(values=("normal", "pageBreakPreview", "pageLayout"))
    topLeftCell = String(allow_none=True)
    colorId = Integer(allow_none=True)
    zoomScale = Integer(allow_none=True)
    zoomScaleNormal = Integer(allow_none=True)
    zoomScaleSheetLayoutView = Integer(allow_none=True)
    zoomScalePageLayoutView = Integer(allow_none=True)
    workbookViewId = Integer()

    def __init__(
            self,
            windowProtection=None,
            showFormulas=None,
            showGridLines=None,
            showRowColHeaders=None,
            showZeros=None,
            rightToLeft=None,
            tabSelected=None,
            showRuler=None,
            showOutlineSymbols=None,
            defaultGridColor=None,
            showWhiteSpace=None,
            view="normal",
            topLeftCell=None,
            colorId=None,
            zoomScale=None,
            zoomScaleNormal=None,
            zoomScaleSheetLayoutView=None,
            zoomScalePageLayoutView=None,
            workbookViewId=None):
        self.windowProtection = windowProtection
        self.showFormulas = showFormulas
        self.showGridLines = showGridLines
        self.showRowColHeaders = showRowColHeaders
        self.showZeros = showZeros
        self.rightToLeft = rightToLeft
        self.tabSelected = tabSelected
        self.showRuler = showRuler
        self.showOutlineSymbols = showOutlineSymbols
        self.defaultGridColor = defaultGridColor
        self.showWhiteSpace = showWhiteSpace
        self.view = view
        self.topLeftCell = topLeftCell
        self.colorId = colorId
        self.zoomScale = zoomScale
        self.zoomScaleNormal = zoomScaleNormal
        self.zoomScaleSheetLayoutView = zoomScaleSheetLayoutView
        self.zoomScalePageLayoutView = zoomScalePageLayoutView
        self.workbookViewId = workbookViewId


class Pane(Serialisable):
    xSplit = Float(allow_none=True)
    ySplit = Float(allow_none=True)
    topLeftCell = String(allow_none=True)
    activePane = Set(values=("bottomRight", "topRight", "bottomLeft", "topLeft"))
    state = Set(values=("split", "frozen", "frozenSplit"))

    def __init__(self,
                 xSplit=None,
                 ySplit=None,
                 topLeftCell=None,
                 activePane="topLeft",
                 state="split"):
        xSplit = xSplit
        ySplit = ySplit
        topLeftCell = topLeftCell
        activePane = activePane
        state = state


class Selection(Serialisable):
    pane = Set(values=("bottomRight", "topRight", "bottomLeft", "topLeft"))
    activeCell = String(allow_none=True)
    activeCellId = Integer(allow_none=True)
    sqref = String(allow_none=True)

    def __init__(self,
                 pane="topLeft",
                 activeCell=None,
                 activeCellId=None,
                 sqref=None):
        pane = pane
        activeCell = activeCell
        activeCellId = activeCellId
        sqref = sqref


class PivotSelection(Serialisable):
    pane = Set(values=("bottomRight", "topRight", "bottomLeft", "topLeft"))
    showHeader = Bool()
    label = Bool()
    data = Bool()
    extendable = Bool()
    count = Integer()
    axis = String(allow_none=True)
    dimension = Integer()
    start = Integer()
    min = Integer()
    max = Integer()
    activeRow = Integer()
    activeCol = Integer()
    previousRow = Integer()
    previousCol = Integer()
    click = Integer()

    def __init__(self,
                 pane=None,
                 showHeader=None,
                 label=None,
                 data=None,
                 extendable=None,
                 count=None,
                 axis=None,
                 dimension=None,
                 start=None,
                 min=None,
                 max=None,
                 activeRow=None,
                 activeCol=None,
                 previousRow=None,
                 previousCol=None,
                 click=None):
        pane = pane
        showHeader = showHeader
        label = label
        data = data
        extendable = extendable
        count = count
        axis = axis
        dimension = dimension
        start = start
        min = min
        max = max
        activeRow = activeRow
        activeCol = activeCol
        previousRow = previousRow
        previousCol = previousCol
        click = click


class PivotArea(Serialisable):

    field = Integer(allow_none=True)
    type = Set(values=())
    dataOnly = Bool()
    labelOnly = Bool()
    grandRow = Bool()
    grandCol = Bool()
    cacheIndex = Bool()
    outline = Bool()
    offset = String()
    collapsedLevelsAreSubtotals = Bool()
    axis = String(allow_none=True)
    fieldPosition = Integer(allow_none=True)

    def __init__(self,
                 field=None,
                 type=None,
                 dataOnly=None,
                 labelOnly=None,
                 grandRow=None,
                 grandCol=None,
                 cacheIndex=None,
                 outline=None,
                 offset=None,
                 collapsedLevelsAreSubtotals=None,
                 axis=None,
                 fieldPosition=None):
        field = field
        type = type
        dataOnly = dataOnly
        labelOnly = labelOnly
        grandRow = grandRow
        grandCol = grandCol
        cacheIndex = cacheIndex
        outline = outline
        offset = offset
        collapsedLevelsAreSubtotals = collapsedLevelsAreSubtotals
        axis = axis
        fieldPosition = fieldPosition


class PivotAreaReferences(Serialisable):

    count = Integer()

    def __init__(self, count=None):
        count = count


class PivotAreaReference(Serialisable):

    field = Integer(allow_none=True)
    count = Integer()
    selected = Bool()
    byPosition = Bool()
    relative = Bool()
    defaultSubtotal = Bool()
    sumSubtotal = Bool()
    countASubtotal = Bool()
    avgSubtotal = Bool()
    maxSubtotal = Bool()
    minSubtotal = Bool()
    productSubtotal = Bool()
    countSubtotal = Bool()
    stdDevSubtotal = Bool()
    stdDevPSubtotal = Bool()
    varSubtotal = Bool()
    varPSubtotal = Bool()

    def __init__(self,
                 field=None,
                 count=None,
                 selected=None,
                 byPosition=None,
                 relative=None,
                 defaultSubtotal=None,
                 sumSubtotal=None,
                 countASubtotal=None,
                 avgSubtotal=None,
                 maxSubtotal=None,
                 minSubtotal=None,
                 productSubtotal=None,
                 countSubtotal=None,
                 stdDevSubtotal=None,
                 stdDevPSubtotal=None,
                 varSubtotal=None,
                 varPSubtotal=None):
        field = field
        count = count
        selected = selected
        byPosition = byPosition
        relative = relative
        defaultSubtotal = defaultSubtotal
        sumSubtotal = sumSubtotal
        countASubtotal = countASubtotal
        avgSubtotal = avgSubtotal
        maxSubtotal = maxSubtotal
        minSubtotal = minSubtotal
        productSubtotal = productSubtotal
        countSubtotal = countSubtotal
        stdDevSubtotal = stdDevSubtotal
        stdDevPSubtotal = stdDevPSubtotal
        varSubtotal = varSubtotal
        varPSubtotal = varPSubtotal


class Index(Serialisable):
    v = Integer()

    def __init__(self, v=None):
        v = v
