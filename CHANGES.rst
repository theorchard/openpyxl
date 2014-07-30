2.0.5 (unreleased)
==================


Bug fixes
---------
348 incorrect casting of boolean strings


2.0.4 (2014-06-25)
==================

Minor changes
-------------
Add a sample file illustrating colours


Bug fixes
---------

#331 DARKYELLOW was incorrect
Correctly handle extend attribute for fonts


2.0.3 (2014-05-22)
==================

Minor changes
-------------

Updated docs


Bug fixes
---------

#319 Cannot load Workbooks with vertAlign styling for fonts


2.0.2 (2014-05-13)
2.0.1 (2014-05-13) # brown bag
2.0.0 (2014-05-13) # brown bag
==================


Major changes
-------------

* This is last release that will support Python 3.2
* Cells are referenced with 1-indexing: A1 == cell(row=1, column=1)
* Use jdcal for more efficient and reliable conversion of datetimes
* Significant speed up when reading files
* Merged immutable styles
* Type inference is disabled by default
* RawCell renamed ReadOnlyCell
* ReadOnlyCell.internal_value and ReadOnlyCell.value now behave the same as Cell
* Provide no size information on unsized worksheets
* Lower memory footprint when reading files


Minor changes
-------------

* All tests converted to pytest
* Pyflakes used for static code analysis
* Sample code in the documentation is automatically run
* Support GradientFills
* BaseColWidth set


Pull requests
-------------
#70 Add filterColumn, sortCondition support to AutoFilter
#80 Reorder worksheets parts
#82 Update API for conditional formatting
#87 Add support for writing Protection styles, others
#89 Better handling of content types when preserving macros


Bug fixes
---------
#46  - ColumnDimension style error
#86 - reader.worksheet.fast_parse sets booleans to integers
#98 - Auto sizing column widths does not work
#137 - Workbooks with chartsheets
#185 - Invalid PageMargins
#230 - Using \v in cells creates invalid files
#243 - IndexError when loading workbook
#263 - Forded conversion of line breaks
#267 - Raise exceptions when passed invalid types
#270 - Cannot open files which use non-standard sheet names or reference Ids
#269 - Handling unsized worksheets in IterableWorksheet
#270 - Handling Workbooks with non-standard references
#275 - Handling auto filters where there are only custom filters
#277 - Harmonise chart and cell coordinates
#280 - Explicit exception raising for invalid characters
#286 - Optimized writer can not handle a datetime.time value
#296 - Cell coordinates not consistent with documentation
#300 - Missing column width causes load_workbook() exception
#304 - Handling Workbooks with absolute paths for worksheets (from Sharepoint)


1.8.6 (2014-05-05)
==================

Minor changes
-------------
Fixed typo for import Elementtree

Bugfixes
--------
#279 Incorrect path for comments files on Windows


1.8.5 (2014-03-25)
==================

Minor changes
-------------
* The '=' string is no longer interpreted as a formula
* When a client writes empty xml tags for cells (e.g. <c r='A1'></c>), reader will not crash


1.8.4 (2014-02-25)
==================

Bugfixes
--------
#260 better handling of undimensioned worksheets
#268 non-ascii in formualae
#282 correct implementation of register_namepsace for Python 2.6


1.8.3 (2014-02-09)
==================

Major changes
-------------
Always parse using cElementTree

Minor changes
-------------
Slight improvements in memory use when parsing

Bugfix #256 - error when trying to read comments with optimised reader
Bugfix #260 - unsized worksheets
Bugfix #264 - only numeric cells can be dates


1.8.2 (2014-01-17)
==================

Bugfix #247 - iterable worksheets open too many files
Bugfix #252 - improved handling of lxml
Bugfix #253 - better handling of unique sheetnames


1.8.1 (2014-01-14)
==================

Bugfix #246


1.8.0 (2014-01-08)
==================

Compatibility
-------------

Support for Python 2.5 dropped.

Major changes
-------------

* Support conditional formatting
* Support lxml as backend
* Support reading and writing comments
* pytest as testrunner now required
* Improvements in charts: new types, more reliable


Minor changes
-------------

* load_workbook now accepts data_only to allow extracting values only from
formulae. Default is false.
* Images can now be anchored to cells
* Docs updated
* Provisional benchmarking
* Added convenience methods for accessing worksheets and cells by key


1.7.0 (2013-10-31)
==================


Major changes
-------------

Drops support for Python < 2.5 and last version to support Python 2.5


Compatibility
-------------

Tests run on Python 2.5, 2.6, 2.7, 3.2, 3.3


Merged pull requests
--------------------

#27 Include more metadata
#41 Able to read files with chart sheets
#45 Configurable Worksheet classes
#3 Correct serialisation of Decimal
#36 Preserve VBA macros when reading files
#44 Handle empty oddheader and oddFooter tags
#43 Fixed issue that the reader never set the active sheet
#33 Reader set value and type explicitly and TYPE_ERROR checking
#22 added page breaks, fixed formula serialization
#39 Fix Python 2.6 compatibility
#47 Improvements in styling


Known bugfixes
--------------

#109
#165
#179
#209
#112
#166
#109
#223
#124
#157


Miscellaneous
-------------

Performance improvements in optimised writer

Docs updated
