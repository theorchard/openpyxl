import cProfile
import os

def read_workbook():
    from openpyxl import load_workbook
    folder = os.path.split(__file__)[0]
    src = os.path.join(folder, "files", "very_large.xlsx")
    wb = load_workbook(src)
    return wb


def rows(wb):
    ws = wb.active
    rows = ws.rows
    for r, row in enumerate(rows):
        for c, col in enumerate(row):
            pass
    print((r+1)* (c+1), "cells")


if __name__ == '__main__':
    wb = read_workbook()
    cProfile.run("rows(wb)", sort="tottime")
