from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

from openpyxl.descriptors import Strict, Bool, Integer, String, Set


class SheetView(Strict):

    """Information about the visible portions of this sheet."""

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
        workbookViewId=None
        ):
        windowProtection = windowProtection
        showFormulas = showFormulas
        showGridLines = showGridLines
        showRowColHeaders = showRowColHeaders
        showZeros = showZeros
        rightToLeft = rightToLeft
        tabSelected = tabSelected
        showRuler = showRuler
        showOutlineSymbols = showOutlineSymbols
        defaultGridColor = defaultGridColor
        showWhiteSpace = showWhiteSpace
        view = view
        topLeftCell = topLeftCell
        colorId = colorId
        zoomScale = zoomScale
        zoomScaleNormal = zoomScaleNormal
        zoomScaleSheetLayoutView = zoomScaleSheetLayoutView
        zoomScalePageLayoutView = zoomScalePageLayoutView
        workbookViewId = workbookViewId
