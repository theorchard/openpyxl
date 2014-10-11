from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

"""Worksheet Properties"""

from openpyxl.compat import safe_string
from openpyxl.descriptors import Strict, String, Bool, Typed
from openpyxl.styles.colors import Color
from openpyxl.xml.constants import SHEET_MAIN_NS
from openpyxl.xml.functions import Element


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
    tabColor = Color(allow_none=True)
    outlinePr = Typed(expected_type=Outline)
    pageSetUpPr = Typed(expected_type=PageSetup)


    def __init__(self,
                 codeName=None,
                 enableFormatConditionsCalculation=None,
                 filterMode=None,
                 published=None,
                 syncHorizontal=None,
                 syncRef=None,
                 syncVertical=None,
                 transitionEvaluation=None,
                 transitionEntry=None,
                 tabColor=None,
                 outlinePr=None,
                 pageSetUpPr=None,
                 ):
        self.codeName = codeName
        self.enableFormatConditionsCalculation = enableFormatConditionsCalculation
        self.filterMode = filterMode
        self.published = published
        self.syncHorizontal = syncHorizontal
        self.syncRef = syncRef
        self.syncVertical = syncVertical
        self.transitionEvaluation = transitionEvaluation
        self.transitionEntry = transitionEntry
        self.tabColor = tabColor
        self.outlinePr = outlinePr
        self.pageSetUpPr = pageSetUpPr


    def __iter__(self):
        for attr in ("codeName", "enableFormatConditionsCalculation",
                     "filterMode", "published", "syncHorizontal", "syncRef",
                     "syncVertical", "transitionEvaluation", "transitionEntry",
                     "tabColor"):
            value = getattr(self, attr)
            if value is not None:
                yield attr, safe_string(value)


def parse_sheetPr(node):
    props = WorksheetProperties(**node.attrib)

    outline = node.find("{%s}outlinePr" % SHEET_MAIN_NS)
    if outline is not None:
        props.outlinePr = Outline(**outline.attrib)

    page_setup = node.find("{%s}pageSetupPr" % SHEET_MAIN_NS)
    if page_setup is not None:
        props.pageSetUpPr = PageSetup(**page_setup.attrib)
    return props


def write_sheetPr(props):
    el = Element("{%s}" % SHEET_MAIN_NS, dict(props))

    outline = props.outlinePr
    if outline:
        el.append(Element("{%s}outlinePr"), dict(outline))

    page_setup = props.pageSetup

    if page_setup:
        el.append(Element("{%s}pageSetupPr"), dict(page_setup))

    return el


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
