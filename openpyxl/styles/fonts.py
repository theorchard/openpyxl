from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl


from openpyxl.descriptors import Float, Integer, Set, Bool, String, Alias, MinMax, NoneSet
from .hashable import HashableObject
from .colors import ColorDescriptor, BLACK


class Font(HashableObject):
    """Font options used in styles."""

    spec = """18.8.22, p.3930"""

    UNDERLINE_DOUBLE = 'double'
    UNDERLINE_DOUBLE_ACCOUNTING = 'doubleAccounting'
    UNDERLINE_SINGLE = 'single'
    UNDERLINE_SINGLE_ACCOUNTING = 'singleAccounting'


    name = String()
    charset = Integer(allow_none=True)
    family = MinMax(min=0, max=14)
    sz = Float()
    size = Alias("sz")
    b = Bool()
    bold = Alias("b")
    i = Bool()
    italic = Alias("i")
    strike = Bool()
    strikethrough = Alias("strike")
    outline = Bool()
    shadow = Bool()
    condense = Bool()
    extend = Bool()
    u = NoneSet(values=(
        UNDERLINE_DOUBLE,
        UNDERLINE_DOUBLE_ACCOUNTING,
        UNDERLINE_SINGLE,
        UNDERLINE_SINGLE_ACCOUNTING
    )
                )
    underline = Alias("u")
    vertAlign = NoneSet(values=('superscript', 'subscript', 'baseline'))
    color = ColorDescriptor()
    scheme = NoneSet(values=("major", "minor"))

    __fields__ = ('name',
                  'sz',
                  'b',
                  'i',
                  'u',
                  'strike',
                  'color',
                  'vertAlign',
                  'charset',
                  'outline',
                  'shadow',
                  'condense',
                  'extend',
                  'family',
                  )

    def __init__(self, name='Calibri', sz=11, b=False, i=False, charset=None,
                 u=None, strike=False, color=BLACK, scheme=None, family=2, size=None,
                 bold=None, italic=None, strikethrough=None, underline=None,
                 vertAlign=None, outline=False, shadow=False, condense=False,
                 extend=False):
        self.name = name
        self.family = family
        if size is not None:
            sz = size
        self.sz = sz
        if bold is not None:
            b = bold
        self.b = b
        if italic is not None:
            i = italic
        self.i = i
        if underline is not None:
            u = underline
        self.u = u
        if strikethrough is not None:
            strike = strikethrough
        self.strike = strike
        self.color = color
        self.vertAlign = vertAlign
        self.charset = charset
        self.outline = outline
        self.shadow = shadow
        self.condense = condense
        self.extend = extend
        self.scheme = scheme


from . colors import Color

DEFAULT_FONT = Font(color=Color(theme=1), scheme="minor")
