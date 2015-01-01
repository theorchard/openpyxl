from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl


import threading
from io import BytesIO

from openpyxl.workbook import Workbook


def test_thread_safe_dump():

    def dump_workbook():
        wb = Workbook(optimized_write=True)
        ws = wb.create_sheet()
        ws.append(range(30))
        wb.save(filename=BytesIO())

    for thread_idx in range(400):
        thread = threading.Thread(target=dump_workbook)
        thread.start()
        print("starting thread %d" % thread_idx)
