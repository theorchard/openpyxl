from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

"""Read shared style definitions"""

# package imports
from openpyxl.compat import OrderedDict, zip
from openpyxl.utils.indexed_list import IndexedList
from openpyxl.styles import (
    numbers,
    Font,
    Fill,
    PatternFill,
    GradientFill,
    Border,
    Side,
    Protection,
    Alignment,
    borders,
)
from openpyxl.formatting.conditional import ConditionaStyle
from openpyxl.styles.colors import COLOR_INDEX, Color
from openpyxl.styles.styleable import StyleId
from openpyxl.styles.named_styles import NamedStyle
from openpyxl.xml.functions import fromstring, safe_iterator, localname
from openpyxl.xml.constants import SHEET_MAIN_NS, ARC_STYLE
from copy import deepcopy


class SharedStylesParser(object):

    def __init__(self, xml_source):
        self.root = fromstring(xml_source)
        self.cell_styles = IndexedList()
        self.cond_styles = []
        self.color_index = COLOR_INDEX
        self.font_list = IndexedList()
        self.fill_list = IndexedList()
        self.border_list = IndexedList()
        self.alignments = IndexedList([Alignment()])
        self.protections = IndexedList([Protection()])
        self.number_formats = IndexedList()

    def parse(self):
        self.parse_custom_num_formats()
        self.parse_color_index()
        self.font_list = IndexedList(self.parse_fonts())
        self.fill_list = IndexedList(self.parse_fills())
        self.border_list = IndexedList(self.parse_borders())
        self.parse_dxfs()
        self.parse_cell_styles()

    def parse_custom_num_formats(self):
        """Read in custom numeric formatting rules from the shared style table"""
        custom_formats = {}
        num_fmts = self.root.findall('{%s}numFmts/{%s}numFmt' % (SHEET_MAIN_NS, SHEET_MAIN_NS))
        for num_fmt_node in num_fmts:
            fmt_code = num_fmt_node.get('formatCode').lower()
            self.number_formats.append(fmt_code)


    def parse_color_index(self):
        """Read in the list of indexed colors"""
        colors =\
            self.root.findall('{%s}colors/{%s}indexedColors/{%s}rgbColor' %
                              (SHEET_MAIN_NS, SHEET_MAIN_NS, SHEET_MAIN_NS))
        if not colors:
            return
        self.color_index = IndexedList([node.get('rgb') for node in colors])


    def parse_dxfs(self):
        """Read in the dxfs effects - used by conditional formatting."""
        for node in self.root.findall("{%s}dxfs/{%s}dxf" % (SHEET_MAIN_NS, SHEET_MAIN_NS) ):
            self.cond_styles.append(ConditionaStyle.from_tree(node))


    def parse_fonts(self):
        """Read in the fonts"""
        fonts = self.root.findall('{%s}fonts/{%s}font' % (SHEET_MAIN_NS, SHEET_MAIN_NS))
        for node in fonts:
            yield Font.from_tree(node)


    def parse_fills(self):
        """Read in the list of fills"""
        fills = self.root.findall('{%s}fills/{%s}fill' % (SHEET_MAIN_NS, SHEET_MAIN_NS))
        for fill in fills:
            yield Fill.from_tree(fill)

    def parse_borders(self):
        """Read in the boarders"""
        borders = self.root.findall('{%s}borders/{%s}border' % (SHEET_MAIN_NS, SHEET_MAIN_NS))
        for border_node in borders:
            yield Border.from_tree(border_node)


    def parse_named_styles(self):
        """
        Extract named styles
        """
        ns = []
        styles_node = self.root.find("{%s}cellStyleXfs" % SHEET_MAIN_NS)
        self._parse_xfs(styles_node)
        _ids = self.cell_styles

        for _name, idx in self._parse_style_names():
            _id = _ids[idx]
            style = NamedStyle(name=_name)
            style.border = self.border_list[_id.border]
            style.fill = self.fill_list[_id.fill]
            style.font = self.font_list[_id.font]
            if _id.alignment:
                style.alignment = self.alignments[_id.alignment]
            if _id.protection:
                style.protection = self.protections[_id.protection]
            ns.append(style)
        self.named_styles = IndexedList(ns)


    def _parse_style_names(self):
        names_node = self.root.find("{%s}cellStyles" % SHEET_MAIN_NS)
        for _name in names_node:
            yield _name.get("name"), int(_name.get("xfId"))


    def parse_cell_styles(self):
        """
        Extract individual cell styles
        """
        node = self.root.find('{%s}cellXfs' % SHEET_MAIN_NS)
        if node is not None:
            self._parse_xfs(node)


    def _parse_xfs(self, node):
        """Read styles from the shared style table"""
        _style_ids = []

        xfs = safe_iterator(node, '{%s}xf' % SHEET_MAIN_NS)
        for xf in xfs:
            attrs = {'alignment':0, 'protection':0}
            d = dict(xf.attrib)

            attrs['font'] = int(d.get('fontId', 0))
            attrs['fill'] = int(d.get('fillId', 0))
            attrs['border'] = int(d.get('borderId', 0))
            attrs['number_format'] = int(d.get('numFmtId', 0))

            al = xf.find('{%s}alignment' % SHEET_MAIN_NS)
            if al is not None:
                alignment = Alignment(**al.attrib)
                attrs['alignment'] = self.alignments.add(alignment)

            prot = xf.find('{%s}protection' % SHEET_MAIN_NS)
            if prot is not None:
                protection = Protection(**prot.attrib)
                attrs['protection'] = self.protections.add(protection)

            _style_ids.append(StyleId(**attrs))
            self.cell_styles = IndexedList(_style_ids)


def read_style_table(archive):
    if ARC_STYLE in archive.namelist():
        xml_source = archive.read(ARC_STYLE)
        p = SharedStylesParser(xml_source)
        p.parse()
        return p


def bool_attrib(element, attr):
    """
    Cast an XML attribute that should be a boolean to a Python equivalent
    None, 'f', '0' and 'false' all cast to False, everything else to true
    """
    value = element.get(attr)
    if not value or value in ("false", "f", "0"):
        return False
    return True
