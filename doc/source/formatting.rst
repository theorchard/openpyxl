Conditional Formatting
======================

There are many types of conditional formatting - below are some examples for setting this within an excel file.

.. :: doctest

>>> from openpyxl import Workbook
>>> from openpyxl.styles import Color, PatternFill, Font, Border
>>> from openpyxl.formatting import ColorScaleRule, CellIsRule, FormulaRule
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
...             ColorScaleRule(start_type='min', start_color=Color('FFAA0000'),
...                           end_type='max', end_color=Color('FF00AA00'))
...                           )
>>>
>>> # Add a three-color scale
>>> ws.conditional_formatting.add('B1:B10',
...                ColorScaleRule(start_type='percentile', start_value=10, start_color=Color('FFAA0000'),
...                            mid_type='percentile', mid_value=50, mid_color=Color('FF0000AA'),
...                            end_type='percentile', end_value=90, end_color=Color('FF00AA00'))
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
>>> # Custom formatting
>>> # There are many types of conditional formatting - it's possible to add additional types directly:
>>> ws.conditional_formatting.add('E1:E10',
...             {'type': 'expression', 'dxf': {'fill': redFill},
...              'formula': ['ISBLANK(E1)'], 'stopIfTrue': '1'})
>>>
>>> # Before writing, call setDxfStyles before saving when adding a conditional format that has a font/border/fill
>>> ws.conditional_formatting.setDxfStyles(wb)
>>> wb.save("test.xlsx")

