from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

"""Write the shared style table."""

# package imports

from openpyxl.compat import safe_string
from openpyxl.collections import IndexedList
from openpyxl.xml.functions import (
    Element,
    SubElement,
    ConditionalElement,
    tostring,
    )
from openpyxl.xml.constants import SHEET_MAIN_NS

from openpyxl.styles import DEFAULTS, Protection
from openpyxl.styles import numbers
from openpyxl.styles.fills import GradientFill, PatternFill


class StyleWriter(object):

    def __init__(self, workbook):
        self.wb = workbook
        self._root = Element('styleSheet', {'xmlns': SHEET_MAIN_NS})

    @property
    def styles(self):
        return self.wb.shared_styles

    def write_table(self):
        number_format_node = SubElement(self._root, 'numFmts')
        fonts_node = self._write_fonts()
        fills_node = self._write_fills()
        borders_node = self._write_borders()
        self._write_cell_style_xfs()
        self._write_cell_xfs(number_format_node, fonts_node, fills_node, borders_node)
        self._write_cell_style()
        self._write_dxfs()
        self._write_table_styles()

        return tostring(self._root)

    def _write_color(self, node, color, key='color'):
        """
        Convert colors encoded as RGB, theme, indexed, auto with tint
        """
        attrs = dict(color)
        SubElement(node, key, attrs)

    def _write_fonts(self):
        """ add fonts part to root
            return {font.crc => index}
        """

        fonts = SubElement(self._root, 'fonts')

        # default
        font_node = SubElement(fonts, 'font')
        SubElement(font_node, 'sz', {'val':'11'})
        SubElement(font_node, 'color', {'theme':'1'})
        SubElement(font_node, 'name', {'val':'Calibri'})
        SubElement(font_node, 'family', {'val':'2'})
        SubElement(font_node, 'scheme', {'val':'minor'})

        return fonts


    def _write_font(self, node, font):
        node = SubElement(node, "font")
        # if present vertAlign has to be at the start otherwise Excel has a problem
        ConditionalElement(node, "vertAlign", font.vertAlign, {'val':font.vertAlign})
        SubElement(node, 'sz', {'val':str(font.size)})
        self._write_color(node, font.color)
        SubElement(node, 'name', {'val':font.name})
        SubElement(node, 'family', {'val': '%d' % font.family})

        # boolean attrs
        for attr in ("b", "i", "outline", "shadow", "condense"):
            ConditionalElement(node, attr, getattr(font, attr))

        # Don't write the 'scheme' element because it appears to prevent
        # the font name from being applied in Excel.
        #SubElement(font_node, 'scheme', {'val':'minor'})

        ConditionalElement(node, "u", font.underline=='single')
        ConditionalElement(node, "charset", font.charset, {'val':str(font.charset)})


    def _write_fills(self):
        fills = SubElement(self._root, 'fills', {'count':'2'})
        fill = SubElement(fills, 'fill')
        SubElement(fill, 'patternFill', {'patternType':'none'})
        fill = SubElement(fills, 'fill')
        SubElement(fill, 'patternFill', {'patternType':'gray125'})
        return fills

    def _write_pattern_fill(self, node, fill):
        if fill != DEFAULTS.fill and fill.fill_type is not None:
            node = SubElement(node, 'patternFill', {'patternType':
                                                    fill.fill_type})
            if fill.start_color != DEFAULTS.fill.start_color:
                self._write_color(node, fill.start_color, 'fgColor')
            if fill.end_color != DEFAULTS.fill.end_color:
                self._write_color(node, fill.end_color, 'bgColor')

    def _write_gradient_fill(self, node, fill):
        node = SubElement(node, 'gradientFill', dict(fill))
        for idx, color in enumerate(fill.stop):
            stop = SubElement(node, "stop", {"position":safe_string(idx)})
            self._write_color(stop, color)

    def _write_borders(self):
        borders = SubElement(self._root, 'borders')

        # default
        border = SubElement(borders, 'border')
        SubElement(border, 'left')
        SubElement(border, 'right')
        SubElement(border, 'top')
        SubElement(border, 'bottom')
        SubElement(border, 'diagonal')
        return borders

    def _write_border(self, node, border):
        """Write the child elements for an individual border section"""
        border_node = SubElement(node, 'border', dict(border))
        for tag, elem in border.children:
            side = SubElement(border_node, tag, dict(elem))
            if elem.color is not None:
                self._write_color(side, elem.color)

    def _write_cell_style_xfs(self):
        cell_style_xfs = SubElement(self._root, 'cellStyleXfs', {'count':'1'})
        xf = SubElement(cell_style_xfs, 'xf',
            {'numFmtId':"0", 'fontId':"0", 'fillId':"0", 'borderId':"0"})

    def _write_cell_xfs(self, number_format_node, fonts_node, fills_node, borders_node):
        """ write styles combinations based on ids found in tables """
        # writing the cellXfs
        cell_xfs = SubElement(self._root, 'cellXfs',
            {'count':'%d' % len(self.styles)})

        # default
        def _get_default_vals():
            return dict(numFmtId='0', fontId='0', fillId='0',
                        xfId='0', borderId='0')

        _fonts = IndexedList()
        _fills = IndexedList()
        _borders = IndexedList()
        _custom_fmts = IndexedList()

        for st in self.styles:
            vals = _get_default_vals()

            font = st.font
            if font != DEFAULTS.font:
                if font not in _fonts:
                    font_id = _fonts.add(font)
                    self._write_font(fonts_node, st.font)
                else:
                    font_id = _fonts.index(font)
                vals['fontId'] = "%d" % (font_id + 1) # There is one default font

                vals['applyFont'] = '1'

            border = st.border
            if st.border != DEFAULTS.border:
                if border not in _borders:
                    border_id = _borders.add(border)
                    self._write_border(borders_node, border)
                else:
                    border_id = _borders.index(border)
                vals['borderId'] = "%d" % (border_id + 1) # There is one default border
                vals['applyBorder'] = '1'


            fill = st.fill
            if fill != DEFAULTS.fill:
                if fill not in _fills:
                    fill_id = _fills.add(st.fill)
                    fill_node = SubElement(fills_node, 'fill')
                    if isinstance(fill, PatternFill):
                        self._write_pattern_fill(fill_node, fill)
                    elif isinstance(fill, GradientFill):
                        self._write_gradient_fill(fill_node, fill)
                else:
                    fill_id = _fills.index(fill)
                vals['fillId'] =  "%d" % (fill_id + 2) # There are two default fills
                vals['applyFill'] = '1'

            nf = st.number_format
            if nf != DEFAULTS.number_format:
                fmt_id = numbers.builtin_format_id(nf)
                if fmt_id is None:
                    if nf not in _custom_fmts:
                        fmt_id = _custom_fmts.add(nf) + 165
                        self._write_number_format(number_format_node, fmt_id, nf)
                    else:
                        fmt_id = _custom_fmts.index(nf)
                vals['numFmtId'] = '%d' % fmt_id
                vals['applyNumberFormat'] = '1'

            if st.alignment != DEFAULTS.alignment:
                vals['applyAlignment'] = '1'

            if st.protection != DEFAULTS.protection:
                vals['applyProtection'] = '1'

            node = SubElement(cell_xfs, 'xf', vals)

            if st.alignment != DEFAULTS.alignment:
                self._write_alignment(node, st.alignment)

            if st.protection != DEFAULTS.protection:
                self._write_protection(node, st.protection)

        fonts_node.attrib["count"] = "%d" % (len(_fonts) + 1)
        borders_node.attrib["count"] = "%d" % (len(_borders) + 1)
        fills_node.attrib["count"] = "%d" % (len(_fills) + 2)
        number_format_node.attrib['count'] = '%d' % len(_custom_fmts)

    def _write_number_format(self, node, fmt_id, format_code):
        fmt_node = SubElement(node, 'numFmt',
                              {'numFmtId':'%d' % (fmt_id),
                               'formatCode':'%s' % format_code}
                              )


    def _write_alignment(self, node, alignment):
        values = dict(alignment)
        if values.get('horizontal', 'general') == 'general':
            del values['horizontal']
        if values.get('vertical', 'bottom') == 'bottom':
            del values['vertical']
        SubElement(node, 'alignment', values)


    def _write_protection(self, node, protection):
        SubElement(node, 'protection', dict(protection))


    def _write_cell_style(self):
        cell_styles = SubElement(self._root, 'cellStyles', {'count':'1'})
        cell_style = SubElement(cell_styles, 'cellStyle',
            {'name':"Normal", 'xfId':"0", 'builtinId':"0"})

    def _write_dxfs(self):
        if self.wb.style_properties and 'dxf_list' in self.wb.style_properties:
            dxfs = SubElement(self._root, 'dxfs', {'count': str(len(self.wb.style_properties['dxf_list']))})
            for d in self.wb.style_properties['dxf_list']:
                dxf = SubElement(dxfs, 'dxf')
                if 'font' in d and d['font'] is not None:
                    font_node = SubElement(dxf, 'font')
                    if d['font'].color is not None:
                        self._write_color(font_node, d['font'].color)
                    ConditionalElement(font_node, 'b', d['font'].bold, 'val')
                    ConditionalElement(font_node, 'i', d['font'].italic, 'val')
                    ConditionalElement(font_node, 'u', d['font'].underline != 'none',
                                       {'val': d['font'].underline})
                    ConditionalElement(font_node, 'strike', d['font'].strikethrough)


                if 'fill' in d:
                    f = d['fill']
                    fill = SubElement(dxf, 'fill')
                    if f.fill_type:
                        node = SubElement(fill, 'patternFill', {'patternType': f.fill_type})
                    else:
                        node = SubElement(fill, 'patternFill')
                    if f.start_color != DEFAULTS.fill.start_color:
                        self._write_color(node, f.start_color, 'fgColor')

                    if f.end_color != DEFAULTS.fill.end_color:
                        self._write_color(node, f.end_color, 'bgColor')

                if 'border' in d:
                    borders = d['border']
                    border = SubElement(dxf, 'border')
                    # caution: respect this order
                    for side in ('left', 'right', 'top', 'bottom'):
                        obj = getattr(borders, side)
                        if obj.border_style is None or obj.border_style == 'none':
                            node = SubElement(border, side)
                        else:
                            node = SubElement(border, side, {'style': obj.border_style})
                            self._write_color(node, obj.color)
        else:
            dxfs = SubElement(self._root, 'dxfs', {'count': '0'})
        return dxfs

    def _write_table_styles(self):

        table_styles = SubElement(self._root, 'tableStyles',
            {'count':'0', 'defaultTableStyle':'TableStyleMedium9',
            'defaultPivotStyle':'PivotStyleLight16'})
