from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

"""Worksheet Properties"""

from openpyxl.compat import safe_string
from openpyxl.descriptors import Strict, String, Bool, Typed
from openpyxl.styles.colors import ARGB


class WorksheetProperties(Strict):

    codeName = String(allow_none=True)
    enableFormatConditionsCalculation = Bool(allow_none=True)
    filterMode = Bool(allow_none=True)
    published = Bool(allow_none=True)
    syncHorizontal = Bool(allow_none=True)
    syncRef = String(allow_none=True)
    syncVertical = Bool(allow_none=True)
    transitionEvaluation = Bool(allow_none=True)
    transitionEntry = Bool(allow_none=True)
    tabColor = ARGB(allow_none=True)
    outlinePr = Typed(expected_type=Outline)
    pageSetUpPr = Typed(expected_type=PageSetup)


    def __init__(self,
                 codeName=None,
                 enableFormatConditionsCalculation=None,
                 ):
        pass


    def __iter__(self):
        pass


class Outline(Strict):

    applyStyles = Bool(allow_none=True)
    summaryBelow = Bool(allow_none=True)
    summaryRight = Bool(allow_none=True)
    showOutlineSymbols = Bool(allow_none=True)


    def __init__(self,
                 applyStyles=None,
                 summaryBelow=None,
                 summaryRight=None,
                 showOutlineSymbols=None
                 ):
        self.applyStyles = applyStyles
        self.summaryBelow = summaryBelow
        self.summaryRight = summaryRight
        self.showOutlineSymbols = showOutlineSymbols


    def __iter__(self):
        for attr in ("applyStyles", "summaryBelow" "summaryRight", "showOutlineSymbols"):
            value = getattr(self, attr)
            if value is not None:
                yield attr, safe_string(value)


class PageSetup(Strict):

    autoPageBreaks = Bool(allow_none=True)
    fitToPage = Bool(allow_none=True)

    def __init__(self, autoPageBreaks=None, fitToPage=None):
        self.autoPageBreaks = autoPageBreaks
        self.fitToPage = fitToPage

    def __iter__(self):
        for attr in ("autoPageBreaks", "fitToPage"):
            value = getattr(self, value)
            if value is not None:
                yielf attr, safe_string(value)
