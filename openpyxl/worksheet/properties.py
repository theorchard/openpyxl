from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

"""Worksheet Properties"""

from openpyxl.compat import safe_string
from openpyxl.descriptors import Strict, String, Bool, Typed
from openpyxl.styles.colors import ColorDescriptor
from openpyxl.xml.constants import SHEET_MAIN_NS
from openpyxl.xml.functions import Element

class Outline(Strict):

    tag = "{%s}outlinePr" % SHEET_MAIN_NS

    applyStyles = Bool(allow_none=True)
    summaryBelow = Bool(allow_none=True)
    summaryRight = Bool(allow_none=True)
    showOutlineSymbols = Bool(allow_none=True)


    def __init__(self,
                 applyStyles=None,
                 summaryBelow=1,
                 summaryRight=1,
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
 
    tag = "{%s}pageSetupPr" % SHEET_MAIN_NS
 
    autoPageBreaks = Bool(allow_none=True)
    fitToPage = Bool(allow_none=True)
 
    def __init__(self, autoPageBreaks=None, fitToPage=None):
        self.autoPageBreaks = autoPageBreaks
        self.fitToPage = fitToPage
 
    def __iter__(self):
        for attr in ("autoPageBreaks", "fitToPage"):
            value = getattr(self, attr)
            if value is not None:
                yield attr, safe_string(value)


class WorksheetProperties(Strict):

    tag = "{%s}sheetPr" % SHEET_MAIN_NS

    codeName = String(allow_none=True)
    enableFormatConditionsCalculation = Bool(allow_none=True)
    filterMode = Bool(allow_none=True)
    published = Bool(allow_none=True)
    syncHorizontal = Bool(allow_none=True)
    syncRef = String(allow_none=True)
    syncVertical = Bool(allow_none=True)
    transitionEvaluation = Bool(allow_none=True)
    transitionEntry = Bool(allow_none=True)
    tabColor = ColorDescriptor(allow_none=True)
    outlinePr = Typed(expected_type=Outline, allow_none=True)
    pageSetUpPr = Typed(expected_type=PageSetup, allow_none=True)


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
                 outlinePr=Outline(),
                 pageSetUpPr=PageSetup(),
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

    outline = node.find(Outline.tag)
    if outline is not None:
        props.outlinePr = Outline(**outline.attrib)

    page_setup = node.find(PageSetup.tag)
    if page_setup is not None:
        props.pageSetUpPr = PageSetup(**page_setup.attrib)
    return props


def write_sheetPr(props):
    el = Element(props.tag, dict(props))

    outline = props.outlinePr
    if outline:
        el.append(Element(outline.tag, dict(outline)))

    page_setup = props.pageSetup

    if page_setup:
        el.append(Element(page_setup.tag, dict(page_setup)))

    return el
