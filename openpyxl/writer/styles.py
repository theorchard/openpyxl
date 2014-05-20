from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

"""Write the shared style table."""

# package imports
from openpyxl.xml.functions import (
    Element,
    SubElement,
    ConditionalElement,
    get_document_content,
    )
from openpyxl.compat import safe_string
from openpyxl.xml.constants import SHEET_MAIN_NS

from openpyxl.styles import DEFAULTS, Protection
from openpyxl.styles.fills import GradientFill, PatternFill


class StyleWriter(object):

    def __init__(self, workbook):
        self.workbook = workbook
        self._style_properties = workbook.style_properties
        self._root = Element('styleSheet', {'xmlns': SHEET_MAIN_NS})

    @property
    def styles(self):
        return self.workbook.shared_styles

    def write_table(self):
        number_format_table = self._write_number_formats()
        fonts_table = self._write_fonts()
        fills_table = self._write_fills()
        borders_table = self._write_borders()
        self._write_cell_style_xfs()
        self._write_cell_xfs(number_format_table, fonts_table, fills_table, borders_table)
        self._write_cell_style()
        self._write_dxfs()
        self._write_table_styles()

        return get_document_content(xml_node=self._root)

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

        # others
        table = {}
        index = 1
        for st in self.styles:
            if st.font != DEFAULTS.font and st.font not in table:
                font_node = SubElement(fonts, 'font')
                table[st.font] = index
                index += 1
                self._write_font(font_node, st.font)

        fonts.attrib["count"] = "%d" % index
        return table


    def _write_font(self, node, font):
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

        table = {}
        index = 2
        for st in self.styles:
            if st.fill != DEFAULTS.fill and st.fill not in table:

                table[st.fill] = index
                node = SubElement(fills, 'fill')
                if isinstance(st.fill, PatternFill):
                    self._write_pattern_fill(node, st.fill)
                elif isinstance(st.fill, GradientFill):
                    self._write_gradient_fill(node, st.fill)
                index += 1

        fills.attrib["count"] = str(index)
        return table

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

        # others
        table = {}
        index = 1
        for st in self.styles:
            if st.border != DEFAULTS.border and st.border not in table:
                table[st.border] = index
                self._write_border(borders, st.border)
                index += 1

        borders.attrib["count"] = str(index)
        return table

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

    def _write_cell_xfs(self, number_format_table, fonts_table, fills_table, borders_table):
        """ write styles combinations based on ids found in tables """

        # writing the cellXfs
        cell_xfs = SubElement(self._root, 'cellXfs',
            {'count':'%d' % (len(self.styles) + 1)})

        # default
        def _get_default_vals():
            return dict(numFmtId='0', fontId='0', fillId='0',
                        xfId='0', borderId='0')

        for st in self.styles:
            vals = _get_default_vals()

            if st.font != DEFAULTS.font:
                vals['fontId'] = str(fonts_table[st.font])
                vals['applyFont'] = '1'

            if st.border != DEFAULTS.border:
                vals['borderId'] = str(borders_table[st.border])
                vals['applyBorder'] = '1'

            if st.fill != DEFAULTS.fill:
                vals['fillId'] = str(fills_table[st.fill])
                vals['applyFill'] = '1'

            if st.number_format != DEFAULTS.number_format:
                vals['numFmtId'] = '%d' % number_format_table[st.number_format]
                vals['applyNumberFormat'] = '1'

            if st.alignment != DEFAULTS.alignment:
                vals['applyAlignment'] = '1'

            if st.protection != DEFAULTS.protection:
                vals['applyProtection'] = '1'

            node = SubElement(cell_xfs, 'xf', vals)

            if st.alignment != DEFAULTS.alignment:
                alignments = {}

                for align_attr in ['horizontal', 'vertical']:
                    if getattr(st.alignment, align_attr) != getattr(DEFAULTS.alignment, align_attr):
                        alignments[align_attr] = getattr(st.alignment, align_attr)

                    if st.alignment.wrap_text != DEFAULTS.alignment.wrap_text:
                        alignments['wrapText'] = '1'

                    if st.alignment.shrink_to_fit != DEFAULTS.alignment.shrink_to_fit:
                        alignments['shrinkToFit'] = '1'

                    if st.alignment.indent > 0:
                        alignments['indent'] = '%s' % st.alignment.indent

                    if st.alignment.text_rotation > 0:
                        alignments['textRotation'] = '%s' % st.alignment.text_rotation
                    elif st.alignment.text_rotation < 0:
                        alignments['textRotation'] = '%s' % (90 - st.alignment.text_rotation)

                SubElement(node, 'alignment', alignments)

            if st.protection != DEFAULTS.protection:
                protections = {}

                if st.protection.locked == Protection.PROTECTION_PROTECTED:
                    protections['locked'] = '1'
                elif st.protection.locked == Protection.PROTECTION_UNPROTECTED:
                    protections['locked'] = '0'

                if st.protection.hidden == Protection.PROTECTION_PROTECTED:
                    protections['hidden'] = '1'
                elif st.protection.hidden == Protection.PROTECTION_UNPROTECTED:
                    protections['hidden'] = '0'

                SubElement(node, 'protection', protections)

    def _write_cell_style(self):
        cell_styles = SubElement(self._root, 'cellStyles', {'count':'1'})
        cell_style = SubElement(cell_styles, 'cellStyle',
            {'name':"Normal", 'xfId':"0", 'builtinId':"0"})

    def _write_dxfs(self):
        if self._style_properties and 'dxf_list' in self._style_properties:
            dxfs = SubElement(self._root, 'dxfs', {'count': str(len(self._style_properties['dxf_list']))})
            for d in self._style_properties['dxf_list']:
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

    def _write_number_formats(self):

        number_format_table = {}

        number_format_list = []
        exceptions_list = []
        num_fmt_id = 165 # start at a greatly higher value as any builtin can go
        num_fmt_offset = 0

        for style in self.styles:

            if not style.number_format in number_format_list  :
                number_format_list.append(style.number_format)

        for number_format in number_format_list:

            if number_format.is_builtin():
                btin = number_format.builtin_format_id(number_format.format_code)
                number_format_table[number_format] = btin
            else:
                number_format_table[number_format] = num_fmt_id + num_fmt_offset
                num_fmt_offset += 1
                exceptions_list.append(number_format)

        num_fmts = SubElement(self._root, 'numFmts',
            {'count':'%d' % len(exceptions_list)})

        for number_format in exceptions_list :
            SubElement(num_fmts, 'numFmt',
                {'numFmtId':'%d' % number_format_table[number_format],
                'formatCode':'%s' % number_format.format_code})

        return number_format_table
