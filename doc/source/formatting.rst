Conditional Formatting
======================

Excel supports three different types of conditional formatting: builtins, standard and custom. Builtins combine specific rules with predefined styles. Standard conditional formats combine specific rules with custom formatting. In additional it is possible to define custom formulae for applying custom formats using differential styles.

.. :: note

The syntax for the different rules varies so much that it is not possible for openpyxl to know whether a rule makes sense or not.


The basic syntax for creating a formatting rule is:

.. :: doctest

>>> from openpyxl.formatting import Rule
>>> from openpyxl.styles import Font, PatternFill, Border
>>> from openpyxl.styles.differential import DifferentialStyle
>>> dxf = DifferentialStyle(font=Font(bold=True), fill=PatternFill(start_color='FFEE1111', end_color='FFEE1111'))
>>> rule = Rule(type='cellIs', dxf=dxf, formula=["10"])

Because the signatures for some rules can be quite verbose there are also some convenience factories for creating them.

Builtin formats
---------------

The builtins conditional formats are:

  * ColorScale
  * IconSet
  * DataBar

Builtin formats contain a sequence of formatting settings which combine a type with an integer for comparison. Possible types are 'num', 'percent', 'max', 'min', 'formula', 'percentile'.


ColorScale
++++++++++

You can have color scales with 2 or 3 colors. 2 color scales produce a gradient from one color to another; 3 color scales use an additional color for 2 gradients.

The full syntax for creating a ColorScale rule is:

.. :: doctest

>>> from openpyxl.formatting.rule import ColorScale, FormatObject
>>> from openpyxl.styles import Color
>>> first = FormatObject(type='min')
>>> last = FormatObject(type='max') 
>>> # colors match the format objects:
>>> colors = [Color('FFAA0000'), Color('FF00AA00')]
>>> cs2 = ColorScale(cfvo=[first, last], color=colors)
>>> # a three color scale would extend the sequences
>>> mid = FormatObject(type='num', val=40) 
>>> colors.insert(1, Color('FF00AA00'))
>>> cs3 = ColorScale(cfvo=[first, mid, last], color=colors)

There is a convenience function for creating ColorScaleRule

.. ::doctest

>>> from openpyxl.formatting.rule import ColorScaleRule
>>> rule = ColorScaleRule(start_type='percentile', start_value=10, start_color='FFAA0000',
...                       mid_type='percentile', mid_value=50, mid_color='FF0000AA',
...                       end_type='percentile', end_value=90, end_color='FF00AA00')

Standard conditional formats
----------------------------

The standard conditional formats are:

  * Average
  * Percent
  * Unique or duplicate
  * Value
  * Rank

.. :: doctest

>>> from openpyxl import Workbook
>>> from openpyxl.styles import Color, PatternFill, Font, Border
>>> from openpyxl.formatting.rule import ColorScaleRule, CellIsRule, FormulaRule
>>>
>>> wb = Workbook()
>>> ws = wb.active
>>>
>>> # Create fill
>>> redFill = PatternFill(start_color='FFEE1111',
...                end_color='FFEE1111',
...                fill_type='solid')
>>>
>>> # Add a two-color scale
>>> # add2ColorScale(range_string, start_type, start_value, start_color, end_type, end_value, end_color)
>>> # Takes colors in excel 'FFRRGGBB' style.
>>> ws.conditional_formatting.add('A1:A10',
...             ColorScaleRule(start_type='min', start_color='FFAA0000',
...                           end_type='max', end_color='FF00AA00')
...                           )
>>>
>>> # Add a three-color scale
>>> ws.conditional_formatting.add('B1:B10',
...                ColorScaleRule(start_type='percentile', start_value=10, start_color='FFAA0000',
...                            mid_type='percentile', mid_value=50, mid_color='FF0000AA',
...                            end_type='percentile', end_value=90, end_color='FF00AA00')
...                              )
>>>
>>> # Add a conditional formatting based on a cell comparison
>>> # addCellIs(range_string, operator, formula, stopIfTrue, wb, font, border, fill)
>>> # Format if cell is less than 'formula'
>>> ws.conditional_formatting.add('C2:C10',
...             CellIsRule(operator='lessThan', formula=['C$1'], stopIfTrue=True, fill=redFill))
>>>
>>> # Format if cell is between 'formula'
>>> ws.conditional_formatting.add('D2:D10',
...             CellIsRule(operator='between', formula=['1','5'], stopIfTrue=True, fill=redFill))
>>>
>>> # Format using a formula
>>> ws.conditional_formatting.add('E1:E10',
...             FormulaRule(formula=['ISBLANK(E1)'], stopIfTrue=True, fill=redFill))
>>>
>>> # Aside from the 2-color and 3-color scales, format rules take fonts, borders and fills for styling:
>>> myFont = Font()
>>> myBorder = Border()
>>> ws.conditional_formatting.add('E1:E10',
...             FormulaRule(formula=['E1=0'], font=myFont, border=myBorder, fill=redFill))
>>>
>>> wb.save("test.xlsx")
