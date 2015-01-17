from __future__ import absolute_import

from openpyxl.descriptors import Integer, String, Typed
from openpyxl.descriptors.serialisable import Serialisable
from openpyxl.styles import (
    Font,
    Fill,
    GradientFill,
    PatternFill,
    Border,
    Alignment,
    Protection,
    )

from openpyxl.xml.functions import localname


class NumFmt(Serialisable):

    numFmtId = Integer()
    formatCode = String()

    def __init__(self,
                 numFmtId=None,
                 formatCode=None,
                ):
        self.numFmtId = numFmtId
        self.formatCode = formatCode


class ConditionalFormat(Serialisable):

    tagname = "dxf"

    __elements__ = ("font", "numFmt", "fill", "alignment", "border", "protection")

    font = Typed(expected_type=Font, allow_none=True)
    numFmt = Typed(expected_type=NumFmt, allow_none=True)
    fill = Typed(expected_type=Fill, allow_none=True)
    alignment = Typed(expected_type=Alignment, allow_none=True)
    border = Typed(expected_type=Border, allow_none=True)
    protection = Typed(expected_type=Protection, allow_none=True)

    def __init__(self,
                 font=None,
                 numFmt=None,
                 fill=None,
                 alignment=None,
                 border=None,
                 protection=None,
                 extLst=None,
                ):
        self.font = font
        self.numFmt = numFmt
        self.fill = fill
        self.alignment = alignment
        self.border = border
        self.protection = protection
        self.extLst = extLst


    @classmethod
    def create(cls, node):
        attrib = {}
        for el in node:
            tag = localname(el)
            if tag == "fill":
                el = [c for c in el][0]
                if "patternFill" in el.tag:
                    obj = PatternFill.create(el)
                elif "gradientFill" in el.tag:
                    obj = GradientFill.create(el)
            else:
                desc = getattr(cls, tag, None)
                obj = desc.expected_type.create(el)
            attrib[tag] = obj
        return cls(**attrib)

