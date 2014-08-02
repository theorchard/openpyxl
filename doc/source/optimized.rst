Optimized reader
================

Sometimes, you will need to open or write extremely large XLSX files,
and the common routines in openpyxl won't be able to handle that load.
Hopefully, there are two modes that enable you to read and write unlimited
amounts of data with (near) constant memory consumption.

Introducing :class:`openpyxl.worksheet.iter_worksheet.IterableWorksheet`::

    from openpyxl import load_workbook
    wb = load_workbook(filename = 'large_file.xlsx', use_iterators = True)
    ws = wb.get_sheet_by_name(name = 'big_data') # ws is now an IterableWorksheet

    for row in ws.iter_rows(): # it brings a new method: iter_rows()

        for cell in row:

            print cell.value

.. warning::

    * :class:`openpyxl.worksheet.iter_worksheet.IterableWorksheet` are read-only
    * range, rows, columns methods and properties are disabled

Cells returned by iter_rows() are not regular :class:`openpyxl.cell.Cell` but
:class:`openpyxl.worksheet.iter_worksheet.RawCell`.

Optimized writer
================

Here again, the regular :class:`openpyxl.worksheet.Worksheet` has been replaced
by a faster alternative, the :class:`openpyxl.writer.dump_worksheet.DumpWorksheet`.
When you want to dump large amounts of data, you might find optimized writer helpful

:: doctest

>>> from openpyxl import Workbook
>>> wb = Workbook(optimized_write = True)
>>> ws = wb.create_sheet()

    # now we'll fill it with 100 rows x 200 columns

>>> for irow in range(100):
...     ws.append(['%d' % i for i in range(200)])

    # wb.save('new_big_file.xlsx') # don't forget to save !

If you want to have cells with styles or comments then use a `WriteOnlyCell`

:: doctest

>>> from openpyxl import Workbook
>>> wb = Workbook(optimized_write = True)
>>> ws = wb.create_sheet()
>>> from openpyxl.writer.dump_worksheet import WriteOnlyCell
>>> from openpyxl.comments import Comment
>>> from openpyxl.styles import Style, Font
>>> cell = WriteOnlyCell(ws, value="hello world")
>>> cell.style = Style(font=Font(name='Courrier', size=36))
>>> cell.comment = Comment(text="A comment", author="Author's Name")


This will append one new row with 3 cells, one text cell with custom font and font size, a float and an empty cell that will be discarded anyway.

.. warning::

    * Those worksheet only have an append() method, it's not possible to access independent cells directly (through cell() or range()). They are write-only.
    * It is able to export unlimited amount of data (even more than Excel can handle actually), while keeping memory usage under 10Mb.
    * A workbook using the optimized writer can only be saved once. After that, every attempt to save the workbook or append() to an existing worksheet will raise an :class:`openpyxl.shared.exc.WorkbookAlreadySaved` exception.
