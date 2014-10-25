from itertools import product
from random import Random
from tempfile import TemporaryFile
import time

import openpyxl
from openpyxl.styles import Style, Alignment, Font


rand = Random()


def generate_all_styles():
    styles = []
    alignments = ['center', 'centerContinuous', 'general', 'justify', 'left',
                  'right']

    font_names = ['Calibri', 'Tahoma', 'Arial', 'Times New Roman']
    font_sizes = range(11, 36, 2)
    bold_options = [True, False]
    underline_options = [True, False]
    italic_options = [True, False]

    for alignment, name, size, bold, underline, italic in product(alignments,
                                                                  font_names,
                                                                  font_sizes,
                                                                  bold_options,
                                                                  underline_options,
                                                                  italic_options):
        s = Style(font=Font(name=name, size=size, italic=italic, underline=underline, bold=bold),
                  alignment=Alignment(horizontal=alignment, vertical=alignment))
        styles.append(s)
    return styles

styles = generate_all_styles()
n = 10000


def optimized_workbook(styles):
    wb = openpyxl.Workbook(optimized_write=True)
    worksheet = wb.create_sheet()
    for _ in range(1, n):
        style = rand.choice(styles)
        worksheet.append([(0, style)])
    return wb


def non_optimized_workbook(styles):
    wb = openpyxl.Workbook()
    for idx in range(1, n):
        worksheet = rand.choice(wb.worksheets)
        cell = worksheet.cell(column=1, row=(idx + 1))
        cell.value = 0
        cell.style = rand.choice(styles)
    return wb


def to_profile(wb, f, n):
    t = -time.time()
    wb.save(f)
    print('took %.4fs for %d styles' % (t + time.time(), n))

for func in (optimized_workbook, non_optimized_workbook):
    print '%s: ' % func.__name__,
    wb = func(styles)
    f = TemporaryFile()
    to_profile(wb, f, n)
