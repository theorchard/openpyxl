Validating cells
================

You can add data validation to a workbook but currently cannot read existing data validation.


Examples
--------

.. :: doctest

>>> from openpyxl import Workbook
>>> from openpyxl.datavalidation import DataValidation, ValidationType
>>>
>>> # Create the workbook and worksheet we'll be working with
>>> wb = Workbook()
>>> ws = wb.active
>>>
>>> # Create a data-validation object with list validation
>>> dv = DataValidation(ValidationType.LIST, formula1='"Dog,Cat,Bat"', allow_blank=True)
>>>
>>> # Optionally set a custom error message
>>> dv.set_error_message('Your entry is not in the list', 'Invalid Entry')
>>>
>>> # Optionally set a custom prompt message
>>> dv.set_prompt_message('Please select from the list', 'List Selection')
>>>
>>> # Add the data-validation object to the worksheet
>>> ws.add_data_validation(dv)

>>> # Create some cells, and add them to the data-validation object
>>> c1 = ws["A1"]
>>> c1.value = "Dog"
>>> dv.add_cell(c1)
>>> c2 = ws["A2"]
>>> c2.value = "An invalid value"
>>> dv.add_cell(c2)
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

    dv = DataValidation(ValidationType.WHOLE)

Any whole number above 100:
::

    dv = DataValidation(ValidationType.WHOLE,
                        ValidationOperator.GREATER_THAN,
                        100)

Any decimal number:
::

    dv = DataValidation(ValidationType.DECIMAL)

Any decimal number between 0 and 1:
::

    dv = DataValidation(ValidationType.DECIMAL,
                        ValidationOperator.BETWEEN,
                        0, 1)

Any date:
::

    dv = DataValidation(ValidationType.DATE)

or time:
::

    dv = DataValidation(ValidationType.TIME)

Any string at most 15 characters:
::

    dv = DataValidation(ValidationType.TEXT_LENGTH,
                        ValidationOperator.LESS_THAN_OR_EQUAL,
                        15)

Custom rule:
::

    dv = DataValidation(ValidationType.CUSTOM,
                        None,
                        "=SOMEFORMULA")

.. note::
    See http://www.contextures.com/xlDataVal07.html for custom rules
