Validating cells
================

You can add data validation to a workbook but currently cannot read existing data validation.


Examples
--------

.. :: doctest

>>> from openpyxl import Workbook
>>> from openpyxl.worksheet.datavalidation import DataValidation, ValidationType
>>>
>>> # Create the workbook and worksheet we'll be working with
>>> wb = Workbook()
>>> ws = wb.active
>>>
>>> # Create a data-validation object with list validation
>>> dv = DataValidation(type="list", formula1='"Dog,Cat,Bat"', allow_blank=True)
>>>
>>> # Optionally set a custom error message
>>> dv.error ='Your entry is not in the list'
>>> dv.errorTitle = 'Invalid Entry'
>>>
>>> # Optionally set a custom prompt message
>>> dv.prompt = 'Please select from the list'
>>> dv.promptTitle = 'List Selection'
>>>
>>> # Add the data-validation object to the worksheet
>>> ws.add_data_validation(dv)

>>> # Create some cells, and add them to the data-validation object
>>> c1 = ws["A1"]
>>> c1.value = "Dog"
>>> dv.add(c1)
>>> c2 = ws["A2"]
>>> c2.value = "An invalid value"
>>> dv.add(c2)
>>>
>>> # Or, apply the validation to a range of cells
>>> dv.ranges.append('B1:B1048576')
>>>
>>> # Write the sheet out.  If you now open the sheet in Excel, you'll find that
>>> # the cells have data-validation applied.
>>> wb.save("test.xlsx")


Other validation examples
-------------------------

Any whole number:
::

    dv = DataValidation(type="whole")

Any whole number above 100:
::

    dv = DataValidation(type="whole",
                        operator="greaterThan",
                        formula1=100)

Any decimal number:
::

    dv = DataValidation(type="decimal")

Any decimal number between 0 and 1:
::

    dv = DataValidation(type="decimal",
                        operator="between",
                        formula1=0,
                        formula2=1)

Any date:
::

    dv = DataValidation(type="date")

or time:
::

    dv = DataValidation(type="time")

Any string at most 15 characters:
::

    dv = DataValidation(type="textLength",
                        operator="lessThanOrEqual"),
                        formula1=15)

Custom rule:
::

    dv = DataValidation(type="custom",
                        formula1"=SOMEFORMULA")

.. note::
    See http://www.contextures.com/xlDataVal07.html for custom rules
