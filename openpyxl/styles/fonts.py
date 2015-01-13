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


    name = String(nested=True)
    charset = Integer(allow_none=True, nested=True)
    family = MinMax(min=0, max=14, nested=True)
    sz = Float(nested=True)
    size = Alias("sz")
    b = Bool(nested=True)
    bold = Alias("b")
    i = Bool(nested=True)
    italic = Alias("i")
    strike = Bool(nested=True)
    strikethrough = Alias("strike")
    outline = Bool(nested=True)
    shadow = Bool(nested=True)
    condense = Bool(nested=True)
    extend = Bool(nested=True)
    u = NoneSet(values=('single', 'double', 'singleAccounting',
                        'doubleAccounting'), nested=True
                )
    underline = Alias("u")
    vertAlign = NoneSet(values=('superscript', 'subscript', 'baseline'), nested=True)
    color = ColorDescriptor()
    scheme = NoneSet(values=("major", "minor"), nested=True)

    tagname = "font"

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

    @classmethod
    def _create_nested(cls, el, tag):
        if tag == "u":
            return el.get("val", "single")
        return super(Font, cls)._create_nested(el, tag)

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
