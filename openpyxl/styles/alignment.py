from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

from openpyxl.descriptors import Set, Float, Bool, Integer, Alias

from .hashable import HashableObject

horizontal_alignments = (
    "general", "left", "center", "right", "fill", "justify", "centerContinuous",
    "distributed"
)
vertical_aligments = (
    "top", "center", "bottom", "justify", "distributed",
)

class Alignment(HashableObject):
    """Alignment options for use in styles."""

    __fields__ = ('horizontal',
                  'vertical',
                  'textRotation',
                  'wrapText',
                  'shrinkToFit',
                  'indent',
                  'relativeIndent',
                  'justifyLastLine',
                  'readingOrder',
                  )
    horizontal = Set(values=horizontal_alignments)
    vertical = Set(values=vertical_aligments)
    textRotation = Integer()
    text_rotation = Alias('textRotation')
    wrapText = Bool()
    wrap_text = Alias('wrapText')
    shrinkToFit = Bool()
    shrink_to_fit = Alias('shrinkToFit')
    indent = Integer()
    relativeIndent = Integer()
    justifyLastLine = Bool()
    readingOrder = Integer()

    def __init__(self, horizontal='general', vertical='bottom',
                 textRotation=0, wrapText=False, shrinkToFit=False, indent=0, relativeIndent=0,
                 justifyLastLine=False, readingOrder=0, text_rotation=None,
                 wrap_text=False, shrink_to_fit=False) :
        self.horizontal = horizontal
        self.vertical = vertical
        self.indent = indent
        self.relativeIndent = relativeIndent
        self.justifyLastLine = justifyLastLine
        self.readingOrder = readingOrder
        if text_rotation is not None:
            textRotation = textRotation
        self.textRotation = textRotation
        if wrap_text is not None:
            wrapText = wrap_text
        self.wrapText = wrapText
        if shrink_to_fit is not None:
            shrinkToFit = shrink_to_fit
        self.shrinkToFit = shrinkToFit
