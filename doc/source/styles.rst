Working with styles
===================

Introduction
------------

Styles are used to change the look of your data while displayed on screen.
They are also used to determine the number format being used for a given cell
or range of cells.

Each :class:`openpyxl.styles.Style` object is composed of many sub-styles, each controlling a
dimension of the styling.

This is what the default `Style` looks like

:: doctest

>>> from openpyxl.styles import Style, PatternFill, Border, Side, Alignment, Protection, Font
>>> s = Style(font=Font(name='Calibri',
...                 size=11,
...                 bold=False,
...                 italic=False,
...                 vertAlign=None,
...                 underline='none',
...                 strike=False,
...                 color='FF000000'),
...       fill=PatternFill(fill_type=None,
...                 start_color='FFFFFFFF',
...                 end_color='FF000000'),
...       border=Border(left=Side(border_style=None,
...                                color='FF000000'),
...                      right=Side(border_style=None,
...                                 color='FF000000'),
...                      top=Side(border_style=None,
...                               color='FF000000'),
...                      bottom=Side(border_style=None,
...                                  color='FF000000'),
...                      diagonal=Side(border_style=None,
...                                    color='FF000000'),
...                      diagonal_direction=0,
...                      outline=Side(border_style=None,
...                                   color='FF000000'),
...                      vertical=Side(border_style=None,
...                                    color='FF000000'),
...                      horizontal=Side(border_style=None,
...                                     color='FF000000')),
...     alignment=Alignment(horizontal='general',
...                         vertical='bottom',
...                         text_rotation=0,
...                         wrap_text=False,
...                         shrink_to_fit=False,
...                         indent=0),
...     number_format='General',
...     protection=Protection(locked='inherit',
...                           hidden='inherit'))
>>>

Pretty big, huh ?
There is one thing to understand about openpyxl's `Styles` : they are immutable.
This means once a `Style` object has been created, it is no longer possible to
alter anything below it.

As you can see from the above box, there is a hierarchy between elements::

    Style > (Font > Color / Fill > Color / Borders > Border > Color / Alignment / NumberFormat / Protection)

So if you want to change the color of a Font, you have to redefine a Style, with a new Font, with a new Color::

>>> from openpyxl.styles import colors
>>> s = Style(font=Font(color=colors.RED))
>>> s.font.color = colors.BLUE # this will not work
>>> blue_s = Style(font=Font(color=colors.BLUE))

However, if you have a Font you want to use multiple times, you are allowed to::

>>> from openpyxl.styles import Font, Color
>>> from openpyxl.styles import colors
>>>
>>> ft = Font(color=colors.RED)
>>> s1 = Style(font=ft, number_format='0%')
>>> s2 = Style(font=ft, number_format='dd-mm-yyyy')


Copying styles
--------------

There is also a `copy()` function, which creates a new style based on another one, by **completely** replacing
sub-elements by others

:: doctest

>>> from openpyxl.styles import Font, Style
>>>
>>> arial = Font(name='Arial', size=14)
>>> tahoma = Font(name='Tahoma')
>>> s1 = Style(font=arial, number_format='0%')
>>> s2 = s1.copy(font=tahoma)
>>> s2.font.name
'Tahoma'
>>> s2.number_format
'0%'
>>> s2.font.size # 11 (tahoma does not re-specify size, so we use the default)
11.0


This might be surprising that we do not use the previous `Font` size,
but this is not a bug, this is because of the immutable nature of styles,
if you want to alter a style, you have to re-define explicitly all the
attributes which are different from the default, even when you copy a `Style`.

Keep this in mind when working with styles and you should be fine.


Basic Font Colors
-----------------
Colors are usually RGB or aRGB hexvalues. The `colors` module contains some constants

:: doctest

>>> from openpyxl.styles import Font
>>> from openpyxl.styles.colors import RED
>>> font = Font(color=RED)
>>> font = Font(color="FFBB00")

There is also support for legacy indexed colors as well as themes and tints

>>> from openpyxl.styles.colors import Color
>>> c = Color(indexed=32)
>>> c = Color(theme=6, tint=0.5)


Applying Styles
---------------
Styles are applied directly to cells

:: doctest

>>> from openpyxl.workbook import Workbook
>>> from openpyxl.styles import Style
>>> wb = Workbook()
>>> ws = wb.active
>>> c = ws['A1']
>>> c.style = Style()

Styles are also applied to columns and rows::

>>> col = ws.column_dimensions['A']
>>> col.style = Style()
>>> row = ws.row_dimensions[1]
>>> row.style = Style()


Edit Print Settings
-------------------
::

>>> from openpyxl.workbook import Workbook
>>>
>>> wb = Workbook()
>>> ws = wb.active
>>>
>>> ws.page_setup.orientation = ws.ORIENTATION_LANDSCAPE
>>> ws.page_setup.paperSize = ws.PAPERSIZE_TABLOID
>>> ws.page_setup.fitToPage = True
>>> ws.page_setup.fitToHeight = 0
>>> ws.page_setup.fitToWidth = 1
>>> ws.page_setup.horizontalCentered = True
>>> ws.page_setup.verticalCentered = True



Header / Footer
---------------
:: doctest

>>> from openpyxl.workbook import Workbook
>>>
>>> wb = Workbook()
>>> ws = wb.worksheets[0]
>>>
>>> ws.header_footer.center_header.text = 'My Excel Page'
>>> ws.header_footer.center_header.font_size = 14
>>> ws.header_footer.center_header.font_name = "Tahoma,Bold"
>>> ws.header_footer.center_header.font_color = "CC3366"

# Or just
>>> ws.header_footer.right_footer.text = 'My Right Footer'


Conditional Formatting
----------------------

There are many types of conditional formatting - below are some examples for setting this within an excel file.

:: doctest

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
