from itertools import product
from random import Random
from tempfile import TemporaryFile
import time

from profilehooks import profile

import openpyxl
from openpyxl.styles import Style, Alignment, Font


rand = Random()


def generate_all_styles():
    styles = []
    alignments = [Alignment.HORIZONTAL_CENTER,
                  Alignment.HORIZONTAL_CENTER_CONTINUOUS,
                  Alignment.HORIZONTAL_GENERAL,
                  Alignment.HORIZONTAL_JUSTIFY,
                  Alignment.HORIZONTAL_LEFT,
                  Alignment.HORIZONTAL_RIGHT]

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
wb = openpyxl.Workbook()
for idx in range(1, n):
    worksheet = rand.choice(wb.worksheets)
    cell = worksheet.cell(column=1, row=(idx + 1))
    cell.value = 0
    cell.style = rand.choice(styles)


# @profile(filename='styles-benchmark.prof')
def to_profile(f, n):
    t = -time.time()
    wb.save(f)
    print 'took %.4fs for %d styles' % (t + time.time(), n)

f = TemporaryFile()
to_profile(f, n)
