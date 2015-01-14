from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

"""Read shared style definitions"""

# package imports
from openpyxl.compat import OrderedDict, zip
from openpyxl.utils.indexed_list import IndexedList
from openpyxl.utils.exceptions import MissingNumberFormat
from openpyxl.styles import (
    Style,
    numbers,
    Font,
    PatternFill,
    GradientFill,
    Border,
    Side,
    Protection,
    Alignment,
    borders,
)
from openpyxl.styles.colors import COLOR_INDEX, Color
from openpyxl.styles.proxy import StyleId
from openpyxl.styles.named_styles import NamedStyle
from openpyxl.xml.functions import fromstring, safe_iterator, localname
from openpyxl.xml.constants import SHEET_MAIN_NS, ARC_STYLE
from copy import deepcopy


class SharedStylesParser(object):

    def __init__(self, xml_source):
        self.root = fromstring(xml_source)
        self.shared_styles = IndexedList()
        self.cell_styles = IndexedList()
        self.cond_styles = []
        self.style_prop = {}
        self.color_index = COLOR_INDEX
        self.font_list = IndexedList()
        self.fill_list = IndexedList()
        self.border_list = IndexedList()
        self.alignments = IndexedList()
        self.protections = IndexedList()

    def parse(self):
        self.parse_custom_num_formats()
        self.parse_color_index()
        self.style_prop['color_index'] = self.color_index
        self.font_list = IndexedList(self.parse_fonts())
        self.fill_list = IndexedList(self.parse_fills())
        self.border_list = IndexedList(self.parse_borders())
        self.parse_dxfs()
        self.parse_cell_styles()

    def parse_custom_num_formats(self):
        """Read in custom numeric formatting rules from the shared style table"""
        custom_formats = {}
        num_fmts = self.root.find('{%s}numFmts' % SHEET_MAIN_NS)
        if num_fmts is not None:
            num_fmt_nodes = safe_iterator(num_fmts, '{%s}numFmt' % SHEET_MAIN_NS)
            for num_fmt_node in num_fmt_nodes:
                fmt_id = int(num_fmt_node.get('numFmtId'))
                fmt_code = num_fmt_node.get('formatCode').lower()
                custom_formats[fmt_id] = fmt_code
        self.custom_num_formats = custom_formats

    def parse_color_index(self):
        """Read in the list of indexed colors"""
        colors = self.root.find('{%s}colors' % SHEET_MAIN_NS)
        if colors is not None:
            indexedColors = colors.find('{%s}indexedColors' % SHEET_MAIN_NS)
            if indexedColors is not None:
                color_nodes = safe_iterator(indexedColors, '{%s}rgbColor' % SHEET_MAIN_NS)
                self.color_index = IndexedList([node.get('rgb') for node in color_nodes])

    def parse_dxfs(self):
        """Read in the dxfs effects - used by conditional formatting."""
        dxf_list = []
        dxfs = self.root.find('{%s}dxfs' % SHEET_MAIN_NS)
        if dxfs is not None:
            nodes = dxfs.findall('{%s}dxf' % SHEET_MAIN_NS)
            for dxf in nodes:
                dxf_item = {}
                font_node = dxf.find('{%s}font' % SHEET_MAIN_NS)
                if font_node is not None:
                    dxf_item['font'] = Font.create(font_node)
                pattern = dxf.find('{%s}fill/{%s}patternFill' % (SHEET_MAIN_NS, SHEET_MAIN_NS))
                if pattern is not None:
                    dxf_item['fill'] = PatternFill.create(pattern)
                gradient = dxf.find('{%s}fill/{%s}gradientFill' % (SHEET_MAIN_NS, SHEET_MAIN_NS))
                if gradient is not None:
                    dxf_item['fill'] = GradientFill.create(gradient)

                border_node = dxf.find('{%s}border' % SHEET_MAIN_NS)
                if border_node is not None:
                    dxf_item['border'] = self.parse_border(border_node)
                dxf_list.append(dxf_item)
        self.cond_styles = dxf_list

    def parse_fonts(self):
        """Read in the fonts"""
        fonts = self.root.find('{%s}fonts' % SHEET_MAIN_NS)
        if fonts is not None:
            for node in safe_iterator(fonts, '{%s}font' % SHEET_MAIN_NS):
                yield Font.create(node)

    def parse_fills(self):
        """Read in the list of fills"""

        fills = self.root.find('{%s}fills' % SHEET_MAIN_NS)
        pattern = '{%s}patternFill' % SHEET_MAIN_NS
        gradient = '{%s}gradientFill' % SHEET_MAIN_NS
        if fills is not None:
            for node in safe_iterator(fills):
                if node.tag == pattern:
                    yield PatternFill.create(node)
                elif node.tag == gradient:
                    yield GradientFill.create(gradient)

    def parse_borders(self):
        """Read in the boarders"""
        borders = self.root.find('{%s}borders' % SHEET_MAIN_NS)
        if borders is not None:
            for border_node in safe_iterator(borders, '{%s}border' % SHEET_MAIN_NS):
                yield self.parse_border(border_node)

    def parse_border(self, border_node):
        """Read individual border"""
        border = dict(border_node.attrib)

        for side in ('left', 'right', 'top', 'bottom', 'diagonal'):
            node = border_node.find('{%s}%s' % (SHEET_MAIN_NS, side))
            if node is not None:
                bside = dict(node.attrib)
                color = node.find('{%s}color' % SHEET_MAIN_NS)
                if color is not None:
                    bside['color'] = Color(**dict(color.attrib))
                border[side] = Side(**bside)
        return Border(**border)


    def parse_named_styles(self):
        """
        Extract named styles
        """
        ns = []
        styles_node = self.root.find("{%s}cellStyleXfs" % SHEET_MAIN_NS)
        _styles, _ids = self._parse_xfs(styles_node)

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
            self.shared_styles, self.cell_styles = self._parse_xfs(node)


    def _parse_xfs(self, node):
        """Read styles from the shared style table"""
        _styles  = []
        _style_ids = []

        builtin_formats = numbers.BUILTIN_FORMATS
        xfs = safe_iterator(node, '{%s}xf' % SHEET_MAIN_NS)
        for index, xf in enumerate(xfs):
            _style = {}

            alignmentId = protectionId = 0
            numFmtId = int(xf.get("numFmtId", 0))
            fontId = int(xf.get("fontId", 0))
            fillId = int(xf.get("fillId", 0))
            borderId = int(xf.get("borderId", 0))

            if numFmtId < 164:
                format_code = builtin_formats.get(numFmtId, 'General')
            else:
                fmt_code = self.custom_num_formats.get(numFmtId)
                if fmt_code is not None:
                    format_code = fmt_code
                else:
                    raise MissingNumberFormat('%s' % numFmtId)
            _style['number_format'] = format_code

            if bool_attrib(xf, 'applyAlignment'):
                al = xf.find('{%s}alignment' % SHEET_MAIN_NS)
                if al is not None:
                    alignment = Alignment(**al.attrib)
                    alignmentId = self.alignments.add(alignment)
                    _style['alignment'] = alignment

            if bool_attrib(xf, 'applyFont'):
                _style['font'] = self.font_list[fontId]

            if bool_attrib(xf, 'applyFill'):
                _style['fill'] = self.fill_list[fillId]

            if bool_attrib(xf, 'applyBorder'):
                _style['border'] = self.border_list[borderId]

            if bool_attrib(xf, 'applyProtection'):
                prot = xf.find('{%s}protection' % SHEET_MAIN_NS)
                if prot is not None:
                    protection = Protection(**prot.attrib)
                    protectionId = self.alignments.add(protection)
                    _style['protection'] = protection

            _styles.append(Style(**_style))
            _style_ids.append(StyleId(alignmentId, borderId, fillId, fontId, numFmtId, protectionId))

        return IndexedList(_styles), IndexedList(_style_ids)


def read_style_table(archive):
    if ARC_STYLE in archive.namelist():
        xml_source = archive.read(ARC_STYLE)
    else:
        return
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
