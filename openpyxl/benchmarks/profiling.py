from io import BytesIO

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
    from openpyxl.writer.lxml_worksheet import write_worksheet_data
    if ws is None:
        ws = make_worksheet()

    out = BytesIO()
    write_worksheet_data(out, ws, [], [])
    #with open("lxl_writer.xml", "wb") as dump:
        #dump.write(out.getvalue())

"""
Sample use
import cProfile
ws = make_worksheet()
cProfile.run("profiling.lxml_writer(ws)", sort="tottime")
"""

if __name__ == '__main__':
    import cProfile
    ws = make_worksheet()
    cProfile.run("sax_writer(ws)", sort="tottime")
    cProfile.run("lxml_writer(ws)", sort="tottime")
