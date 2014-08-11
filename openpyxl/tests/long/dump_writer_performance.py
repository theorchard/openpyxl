import openpyxl
from openpyxl.compat import range

import tempfile

import pytest

@pytest.mark.parametrize("mode", (True, False))
def test_large_append(mode):
    print("Using write only mode {0}".format(mode))
    wb = openpyxl.Workbook(optimized_write=mode)
    ws = wb.create_sheet()
    row = ('this is some text', 3.14)
    total_rows = int(2e4)
    for idx in range(total_rows):
        if not idx % 10000:
            print("%.2f%%" % (100 * (float(idx) / float(total_rows))))
        ws.append(row)
    wb.save(tempfile.TemporaryFile(mode='wb'))
