from io import BytesIO
from lxml.etree import xmlfile
from random import randint

from openpyxl import Workbook
from openpyxl.xml.functions import XMLGenerator

def make_worksheet():
    wb = Workbook()
    ws = wb.active
    for i in range(1000):
        ws.append(list(range(100)))
    return ws


def sax_writer(ws=None):
    from openpyxl.writer.worksheet import write_worksheet_data
    if ws is None:
        ws = make_worksheet()
    out = BytesIO()
    doc = XMLGenerator(out)
    write_worksheet_data(doc, ws, [], [])
    #with open("sax_writer.xml", "wb") as dump:
        #dump.write(out.getvalue())


def lxml_writer(ws=None):
    from openpyxl.writer.lxml_worksheet import write_rows
    if ws is None:
        ws = make_worksheet()

    out = BytesIO()
    with xmlfile(out) as xf:
        write_rows(xf, ws, [], [])
    #with open("lxml_writer.xml", "wb") as dump:
        #dump.write(out.getvalue())
    #ws.parent.save("lxml_writer.xlsx")


COLUMNS = 100
ROWS = 1000
BOLD = 1
ITALIC = 2
UNDERLINE = 4
RED_BG = 8
formatData = [[None] * COLUMNS for _ in range(ROWS)]

def generate_format_data():
    for row in range(ROWS):
        for col in range(COLUMNS):
            formatData[row][col] = randint(1, 15)


def styled_sheet():
    from openpyxl import Workbook
    from openpyxl.styles import Font, Style, PatternFill, Color, colors

    wb = Workbook()
    ws = wb.active
    ws.title = 'Test 1'

    red_fill = PatternFill(fill_type='solid', fgColor=Color(colors.RED), bgColor=Color(colors.RED))
    empty_fill = PatternFill()
    styles = []
    # pregenerate relevant styles
    for row in range(ROWS):
        _row = []
        for col in range(COLUMNS):
            cell = ws.cell(row=row+1, column=col+1)
            cell.value = 1
            font = {}
            fill = PatternFill()
            if formatData[row][col] & BOLD:
                font['bold'] = True
            if formatData[row][col] & ITALIC:
                font['italic'] = True
            if formatData[row][col] & UNDERLINE:
                font['underline'] = 'single'
            if formatData[row][col] & RED_BG:
                fill = red_fill
            cell.style = Style(font=Font(**font), fill=fill)

    #wb.save(get_output_path('test_openpyxl_style_std_pregen.xlsx'))


"""
Sample use
import cProfile
ws = make_worksheet()
cProfile.run("profiling.lxml_writer(ws)", sort="tottime")
"""

if __name__ == '__main__':
    import cProfile
    ws = make_worksheet()
    cProfile.run("make_worksheet()", sort="tottime")
    #cProfile.run("sax_writer(ws)", sort="tottime")
    #cProfile.run("lxml_writer(ws)", sort="tottime")
    #generate_format_data()
    #cProfile.run("styled_sheet()", sort="tottime")
