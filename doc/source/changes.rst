2.2.0 (unreleased)
==================


Major changes
-------------
* Cell styles deprecated, use formatting objects (fonts, fills, borders, etc.) directly instead
* Charts will no longer try and calculate axes by default
* Support for template file types - PR21
* Moved ancillary functions and classes into utils package - single place of reference
* `PR 34 <https://bitbucket.org/openpyxl/openpyxl/pull-request/34/>`_ Fully support page setup
* Removed SAX-based XML Generator. Special thanks to Elias Rabel for implementing xmlfile for xml.etree
* Preserve sheet view definitions in existing files (frozen panes, zoom, etc.)


Bug fixes
---------
* `#103 <https://bitbucket.org/openpyxl/openpyxl/issue/103>`_ Set the zoom of a sheet
* `#199 <https://bitbucket.org/openpyxl/openpyxl/issue/199>`_ Hide gridlines
* `#215 <https://bitbucket.org/openpyxl/openpyxl/issue/215>`_ Preserve sheet view setings
* `#262 <https://bitbucket.org/openpyxl/openpyxl/issue/262>`_ Set the zoom of a sheet
* `#392 <https://bitbucket.org/openpyxl/openpyxl/issue/392>`_ Worksheet header not read
* `#387 <https://bitbucket.org/openpyxl/openpyxl/issue/387>`_ Cannot read files without styles.xml
* `#410 <https://bitbucket.org/openpyxl/openpyxl/issue/410>`_ Exception when preserving whitespace in strings
* `#417 <https://bitbucket.org/openpyxl/openpyxl/issue/417>`_ Cannot create print titles
* `#420 <https://bitbucket.org/openpyxl/openpyxl/issue/420>`_ Rename confusing constants
* `#422 <https://bitbucket.org/openpyxl/openpyxl/issue/422>`_ Preserve color index in a workbook if it differs from the standard


Minor changes
-------------
* Use a 2-way cache for column index lookups
* Clean up tests in cells
* `PR 40 <https://bitbucket.org/openpyxl/openpyxl/pull-request/40/>`_ Support frozen panes and autofilter in write-only mode
* Use ws.calculate_dimension(force=True) in read-only mode for unsized worksheets


2.1.5 (2015-02-18)
==================


Bug fixes
---------
* `#403 <https://bitbucket.org/openpyxl/openpyxl/issue/403>`_ Cannot add comments in write-only mode
* `#401 <https://bitbucket.org/openpyxl/openpyxl/issue/401>`_ Creating cells in an empty row raises an exception
* `#408 <https://bitbucket.org/openpyxl/openpyxl/issue/408>`_ from_excel adjustment for Julian dates 1 < x < 60
* `#409 <https://bitbucket.org/openpyxl/openpyxl/issue/409>`_ refersTo is an optional attribute


Minor changes
-------------
* Allow cells to be appended to standard worksheets for code compatibility with write-only mode.


2.1.4 (2014-12-16)
==================


Bug fixes
---------

* `#393 <https://bitbucket.org/openpyxl/openpyxl/issue/393>`_ IterableWorksheet skips empty cells in rows
* `#394 <https://bitbucket.org/openpyxl/openpyxl/issue/394>`_ Date format is applied to all columns (while only first column contains dates)
* `#395 <https://bitbucket.org/openpyxl/openpyxl/issue/395>`_ temporary files not cleaned properly
* `#396 <https://bitbucket.org/openpyxl/openpyxl/issue/396>`_ Cannot write "=" in Excel file
* `#398 <https://bitbucket.org/openpyxl/openpyxl/issue/398>`_ Cannot write empty rows in write-only mode with LXML installed


Minor changes
-------------
* Add relation namespace to root element for compatibility with iWork
* Serialize comments relation in LXML-backend


2.1.3 (2014-12-09)
==================


Minor changes
-------------
* `PR 31 <https://bitbucket.org/openpyxl/openpyxl/pull-request/31/>`_ Correct tutorial
* `PR 32 <https://bitbucket.org/openpyxl/openpyxl/pull-request/32/>`_ See #380
* `PR 37 <https://bitbucket.org/openpyxl/openpyxl/pull-request/37/>`_ Bind worksheet to ColumnDimension objects


Bug fixes
---------
* `#379 <https://bitbucket.org/openpyxl/openpyxl/issue/379>`_ ws.append() doesn't set RowDimension Correctly
* `#380 <https://bitbucket.org/openpyxl/openpyxl/issue/379>`_ empty cells formatted as datetimes raise exceptions


2.1.2 (2014-10-23)
==================


Minor changes
-------------
* `PR 30 <https://bitbucket.org/openpyxl/openpyxl/pull-request/30/>`_ Fix regex for positive exponentials
* `PR 28 <https://bitbucket.org/openpyxl/openpyxl/pull-request/28/>`_ Fix for #328


Bug fixes
---------
* `#120 <https://bitbucket.org/openpyxl/openpyxl/issue/120>`_, `#168 <https://bitbucket.org/openpyxl/openpyxl/issue/168>`_ defined names with formulae raise exceptions, `#292 <https://bitbucket.org/openpyxl/openpyxl/issue/292>`_
* `#328 <https://bitbucket.org/openpyxl/openpyxl/issue/328/>`_ ValueError when reading cells with hyperlinks
* `#369 <https://bitbucket.org/openpyxl/openpyxl/issue/369>`_ IndexError when reading definedNames
* `#372 <https://bitbucket.org/openpyxl/openpyxl/issue/372>`_ number_format not consistently applied from styles


2.1.1 (2014-10-08)
==================


Minor changes
-------------
* PR 20 Support different workbook code names
* Allow auto_axis keyword for ScatterCharts


Bug fixes
---------

* `#332 <https://bitbucket.org/openpyxl/openpyxl/issue/332>`_ Fills lost in ConditionalFormatting
* `#360 <https://bitbucket.org/openpyxl/openpyxl/issue/360>`_ Support value="none" in attributes
* `#363 <https://bitbucket.org/openpyxl/openpyxl/issue/363>`_ Support undocumented value for textRotation
* `#364 <https://bitbucket.org/openpyxl/openpyxl/issue/364>`_ Preserve integers in read-only mode
* `#366 <https://bitbucket.org/openpyxl/openpyxl/issue/366>`_ Complete read support for DataValidation
* `#367 <https://bitbucket.org/openpyxl/openpyxl/issue/367>`_ Iterate over unsized worksheets


2.1.0 (2014-09-21)
==================

Major changes
-------------
* "read_only" and "write_only" new flags for workbooks
* Support for reading and writing worksheet protection
* Support for reading hidden rows
* Cells now manage their styles directly
* ColumnDimension and RowDimension object manage their styles directly
* Use xmlfile for writing worksheets if available - around 3 times faster
* Datavalidation now part of the worksheet package


Minor changes
-------------
* Number formats are now just strings
* Strings can be used for RGB and aRGB colours for Fonts, Fills and Borders
* Create all style tags in a single pass
* Performance improvement when appending rows
* Cleaner conversion of Python to Excel values
* PR6 reserve formatting for empty rows
* standard worksheets can append from ranges and generators


Bug fixes
---------
* `#153 <https://bitbucket.org/openpyxl/openpyxl/issue/153>`_ Cannot read visibility of sheets and rows
* `#181 <https://bitbucket.org/openpyxl/openpyxl/issue/181>`_ No content type for worksheets
* `241 <https://bitbucket.org/openpyxl/openpyxl/issue/241>`_ Cannot read sheets with inline strings
* `322 <https://bitbucket.org/openpyxl/openpyxl/issue/322>`_ 1-indexing for merged cells
* `339 <https://bitbucket.org/openpyxl/openpyxl/issue/339>`_ Correctly handle removal of cell protection
* `341 <https://bitbucket.org/openpyxl/openpyxl/issue/341>`_ Cells with formulae do not round-trip
* `347 <https://bitbucket.org/openpyxl/openpyxl/issue/347>`_ Read DataValidations
* `353 <https://bitbucket.org/openpyxl/openpyxl/issue/353>`_ Support Defined Named Ranges to external workbooks


2.0.5 (2014-08-08)
==================


Bug fixes
---------
* `#348 <https://bitbucket.org/openpyxl/openpyxl/issue/348>`_ incorrect casting of boolean strings
* `#349 <https://bitbucket.org/openpyxl/openpyxl/issue/349>`_ roundtripping cells with formulae


2.0.4 (2014-06-25)
==================

Minor changes
-------------
* Add a sample file illustrating colours


Bug fixes
---------

* `#331 <https://bitbucket.org/openpyxl/openpyxl/issue/331>`_ DARKYELLOW was incorrect
* Correctly handle extend attribute for fonts


2.0.3 (2014-05-22)
==================

Minor changes
-------------

* Updated docs


Bug fixes
---------

* `#319 <https://bitbucket.org/openpyxl/openpyxl/issue/319>`_ Cannot load Workbooks with vertAlign styling for fonts


2.0.2 (2014-05-13)
==================

2.0.1 (2014-05-13)  brown bag
=============================

2.0.0 (2014-05-13)  brown bag
=============================


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
* #70 Add filterColumn, sortCondition support to AutoFilter
* #80 Reorder worksheets parts
* #82 Update API for conditional formatting
* #87 Add support for writing Protection styles, others
* #89 Better handling of content types when preserving macros


Bug fixes
---------
* `#46 <https://bitbucket.org/openpyxl/openpyxl/issue/46>`_ ColumnDimension style error
* `#86 <https://bitbucket.org/openpyxl/openpyxl/issue/86>`_ reader.worksheet.fast_parse sets booleans to integers
* `#98 <https://bitbucket.org/openpyxl/openpyxl/issue/98>`_ Auto sizing column widths does not work
* `#137 <https://bitbucket.org/openpyxl/openpyxl/issue/137>`_ Workbooks with chartsheets
* `#185 <https://bitbucket.org/openpyxl/openpyxl/issue/185>`_  Invalid PageMargins
* `#230 <https://bitbucket.org/openpyxl/openpyxl/issue/230>`_ Using \v in cells creates invalid files
* `#243 <https://bitbucket.org/openpyxl/openpyxl/issue/243>`_ - IndexError when loading workbook
* `#263 <https://bitbucket.org/openpyxl/openpyxl/issue/263>`_ - Forded conversion of line breaks
* `#267 <https://bitbucket.org/openpyxl/openpyxl/issue/267>`_ - Raise exceptions when passed invalid types
* `#270 <https://bitbucket.org/openpyxl/openpyxl/issue/270>`_ - Cannot open files which use non-standard sheet names or reference Ids
* `#269 <https://bitbucket.org/openpyxl/openpyxl/issue/269>`_ - Handling unsized worksheets in IterableWorksheet
* `#270 <https://bitbucket.org/openpyxl/openpyxl/issue/270>`_ - Handling Workbooks with non-standard references
* `#275 <https://bitbucket.org/openpyxl/openpyxl/issue/275>`_ - Handling auto filters where there are only custom filters
* `#277 <https://bitbucket.org/openpyxl/openpyxl/issue/277>`_ - Harmonise chart and cell coordinates
* `#280 <https://bitbucket.org/openpyxl/openpyxl/issue/280>`_- Explicit exception raising for invalid characters
* `#286 <https://bitbucket.org/openpyxl/openpyxl/issue/286>`_ - Optimized writer can not handle a datetime.time value
* `#296 <https://bitbucket.org/openpyxl/openpyxl/issue/296>`_ - Cell coordinates not consistent with documentation
* `#300 <https://bitbucket.org/openpyxl/openpyxl/issue/300>`_ - Missing column width causes load_workbook() exception
* `#304 <https://bitbucket.org/openpyxl/openpyxl/issue/304>`_ - Handling Workbooks with absolute paths for worksheets (from Sharepoint)


1.8.6 (2014-05-05)
==================

Minor changes
-------------
Fixed typo for import Elementtree

Bugfixes
--------
* `#279 <https://bitbucket.org/openpyxl/openpyxl/issue/279>`_ Incorrect path for comments files on Windows


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
* `#260 <https://bitbucket.org/openpyxl/openpyxl/issue/260>`_ better handling of undimensioned worksheets
* `#268 <https://bitbucket.org/openpyxl/openpyxl/issue/268>`_ non-ascii in formualae
* `#282 <https://bitbucket.org/openpyxl/openpyxl/issue/282>`_ correct implementation of register_namepsace for Python 2.6


1.8.3 (2014-02-09)
==================

Major changes
-------------
Always parse using cElementTree

Minor changes
-------------
Slight improvements in memory use when parsing

* `#256 <https://bitbucket.org/openpyxl/openpyxl/issue/256>`_ - error when trying to read comments with optimised reader
* `#260 <https://bitbucket.org/openpyxl/openpyxl/issue/260>`_ - unsized worksheets
* `#264 <https://bitbucket.org/openpyxl/openpyxl/issue/264>`_ - only numeric cells can be dates


1.8.2 (2014-01-17)
==================

* `#247 <https://bitbucket.org/openpyxl/openpyxl/issue/247>`_ - iterable worksheets open too many files
* `#252 <https://bitbucket.org/openpyxl/openpyxl/issue/252>`_ - improved handling of lxml
* `#253 <https://bitbucket.org/openpyxl/openpyxl/issue/253>`_ - better handling of unique sheetnames


1.8.1 (2014-01-14)
==================

* `#246 <https://bitbucket.org/openpyxl/openpyxl/issue/246>`_


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

* 27 Include more metadata
* 41 Able to read files with chart sheets
* 45 Configurable Worksheet classes
* 3 Correct serialisation of Decimal
* 36 Preserve VBA macros when reading files
* 44 Handle empty oddheader and oddFooter tags
* 43 Fixed issue that the reader never set the active sheet
* 33 Reader set value and type explicitly and TYPE_ERROR checking
* 22 added page breaks, fixed formula serialization
* 39 Fix Python 2.6 compatibility
* 47 Improvements in styling


Known bugfixes
--------------

* `#109 <https://bitbucket.org/openpyxl/openpyxl/issue/109>`_
* `#165 <https://bitbucket.org/openpyxl/openpyxl/issue/165>`_
* `#179 <https://bitbucket.org/openpyxl/openpyxl/issue/179>`_
* `#209 <https://bitbucket.org/openpyxl/openpyxl/issue/209>`_
* `#112 <https://bitbucket.org/openpyxl/openpyxl/issue/112>`_
* `#166 <https://bitbucket.org/openpyxl/openpyxl/issue/166>`_
* `#109 <https://bitbucket.org/openpyxl/openpyxl/issue/109>`_
* `#223 <https://bitbucket.org/openpyxl/openpyxl/issue/223>`_
* `#124 <https://bitbucket.org/openpyxl/openpyxl/issue/124>`_
* `#157 <https://bitbucket.org/openpyxl/openpyxl/issue/157>`_


Miscellaneous
-------------

Performance improvements in optimised writer

Docs updated
