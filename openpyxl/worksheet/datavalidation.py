from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl

from itertools import groupby, chain

from openpyxl.compat import OrderedDict, safe_string
from openpyxl.cell import coordinate_from_string
from openpyxl.worksheet import cells_from_range
from openpyxl.xml.constants import SHEET_MAIN_NS
from openpyxl.xml.functions import Element, safe_iterator


def collapse_cell_addresses(cells, input_ranges=()):
    """ Collapse a collection of cell co-ordinates down into an optimal
        range or collection of ranges.

        E.g. Cells A1, A2, A3, B1, B2 and B3 should have the data-validation
        object applied, attempt to collapse down to a single range, A1:B3.

        Currently only collapsing contiguous vertical ranges (i.e. above
        example results in A1:A3 B1:B3).  More work to come.
    """
    keyfunc = lambda x: x[0]

    # Get the raw coordinates for each cell given
    raw_coords = [coordinate_from_string(cell) for cell in cells]

    # Group up as {column: [list of rows]}
    grouped_coords = OrderedDict((k, [c[1] for c in g]) for k, g in
                          groupby(sorted(raw_coords, key=keyfunc), keyfunc))
    ranges = list(input_ranges)

    # For each column, find contiguous ranges of rows
    for column in grouped_coords:
        rows = sorted(grouped_coords[column])
        grouped_rows = [[r[1] for r in list(g)] for k, g in
                        groupby(enumerate(rows),
                        lambda x: x[0] - x[1])]
        for rows in grouped_rows:
            if len(rows) == 0:
                pass
            elif len(rows) == 1:
                ranges.append("%s%d" % (column, rows[0]))
            else:
                ranges.append("%s%d:%s%d" % (column, rows[0], column, rows[-1]))

    return " ".join(ranges)


def expand_cell_ranges(range_string):
    """
    Expand cell ranges to a sequence of addresses.
    Reverse of collapse_cell_addresses
    Eg. converts "A1:A2 B1:B2" to (A1, A2, B1, B2)
    """
    cells = []
    for rs in range_string.split():
        cells.extend(cells_from_range(rs))
    return list(chain.from_iterable(cells))


"""
  <xsd:complexType name="CT_DataValidations">
    <xsd:sequence>
      <xsd:element name="dataValidation" type="CT_DataValidation" minOccurs="1"
        maxOccurs="unbounded"/>
    </xsd:sequence>
    <xsd:attribute name="disablePrompts" type="xsd:boolean" use="optional" default="false"/>
    <xsd:attribute name="xWindow" type="xsd:unsignedInt" use="optional"/>
    <xsd:attribute name="yWindow" type="xsd:unsignedInt" use="optional"/>
    <xsd:attribute name="count" type="xsd:unsignedInt" use="optional"/>
  </xsd:complexType>
  <xsd:complexType name="CT_DataValidation">
    <xsd:sequence>
      <xsd:element name="formula1" type="ST_Formula" minOccurs="0" maxOccurs="1"/>
      <xsd:element name="formula2" type="ST_Formula" minOccurs="0" maxOccurs="1"/>
    </xsd:sequence>
    <xsd:attribute name="type" type="ST_DataValidationType" use="optional" default="none"/>
    <xsd:attribute name="errorStyle" type="ST_DataValidationErrorStyle" use="optional"
      default="stop"/>
    <xsd:attribute name="imeMode" type="ST_DataValidationImeMode" use="optional" default="noControl"/>
    <xsd:attribute name="operator" type="ST_DataValidationOperator" use="optional" default="between"/>
    <xsd:attribute name="allowBlank" type="xsd:boolean" use="optional" default="false"/>
    <xsd:attribute name="showDropDown" type="xsd:boolean" use="optional" default="false"/>
    <xsd:attribute name="showInputMessage" type="xsd:boolean" use="optional" default="false"/>
    <xsd:attribute name="showErrorMessage" type="xsd:boolean" use="optional" default="false"/>
    <xsd:attribute name="errorTitle" type="s:ST_Xstring" use="optional"/>
    <xsd:attribute name="error" type="s:ST_Xstring" use="optional"/>
    <xsd:attribute name="promptTitle" type="s:ST_Xstring" use="optional"/>
    <xsd:attribute name="prompt" type="s:ST_Xstring" use="optional"/>
    <xsd:attribute name="sqref" type="ST_Sqref" use="required"/>
  </xsd:complexType>
  <xsd:simpleType name="ST_DataValidationType">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="none"/>
      <xsd:enumeration value="whole"/>
      <xsd:enumeration value="decimal"/>
      <xsd:enumeration value="list"/>
      <xsd:enumeration value="date"/>
      <xsd:enumeration value="time"/>
      <xsd:enumeration value="textLength"/>
      <xsd:enumeration value="custom"/>
    </xsd:restriction>
  </xsd:simpleType>
  <xsd:simpleType name="ST_DataValidationOperator">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="between"/>
      <xsd:enumeration value="notBetween"/>
      <xsd:enumeration value="equal"/>
      <xsd:enumeration value="notEqual"/>
      <xsd:enumeration value="lessThan"/>
      <xsd:enumeration value="lessThanOrEqual"/>
      <xsd:enumeration value="greaterThan"/>
      <xsd:enumeration value="greaterThanOrEqual"/>
    </xsd:restriction>
  </xsd:simpleType>
  <xsd:simpleType name="ST_DataValidationErrorStyle">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="stop"/>
      <xsd:enumeration value="warning"/>
      <xsd:enumeration value="information"/>
    </xsd:restriction>
  </xsd:simpleType>
  <xsd:simpleType name="ST_DataValidationImeMode">
    <xsd:restriction base="xsd:string">
      <xsd:enumeration value="noControl"/>
      <xsd:enumeration value="off"/>
      <xsd:enumeration value="on"/>
      <xsd:enumeration value="disabled"/>
      <xsd:enumeration value="hiragana"/>
      <xsd:enumeration value="fullKatakana"/>
      <xsd:enumeration value="halfKatakana"/>
      <xsd:enumeration value="fullAlpha"/>
      <xsd:enumeration value="halfAlpha"/>
      <xsd:enumeration value="fullHangul"/>
      <xsd:enumeration value="halfHangul"/>
    </xsd:restriction>
  </xsd:simpleType>
"""


class DataValidation(object):


    error = None
    errorTitle = None
    prompt = None
    promptTitle = None

    def __init__(self,
                 validation_type=None, # remove in future
                 operator=None,
                 formula1=None,
                 formula2=None,
                 allow_blank=False,
                 showErrorMessage=True,
                 showInputMessage=True,
                 allowBlank=None,
                 type=None,
                 sqref=None):

        self.type = validation_type
        self.operator = operator
        if formula1 is not None:
            self.formula1 = str(formula1)
        if formula2 is not None:
            self.formula2 = str(formula2)
        self.allowBlank = allow_blank
        if allowBlank is not None:
            self.allowBlank = allowBlank
        self.showErrorMessage = showErrorMessage
        self.showInputMessage = showInputMessage
        if type is not None:
            self.type = type
        self.cells = []
        self.ranges = []
        if sqref is not None:
            self.sqref = sqref

    def add_cell(self, cell):
        """Adds a openpyxl.cell to this validator"""
        self.cells.append(cell.coordinate)

    def set_error_message(self, error, error_title="Validation Error"):
        """Creates a custom error message, displayed when a user changes a cell
           to an invalid value"""
        self.errorTitle = error_title
        self.error = error

    def set_prompt_message(self, prompt, prompt_title="Validation Prompt"):
        """Creates a custom prompt message"""
        self.promptTitle = prompt_title
        self.prompt = prompt

    @property
    def sqref(self):
        return collapse_cell_addresses(self.cells, self.ranges)

    @sqref.setter
    def sqref(self, range_string):
        self.cells = expand_cell_ranges(range_string)

    def __iter__(self):
        for attr in ('type', 'allowBlank', 'operator', 'sqref',
                     'showInputMessage', 'showErrorMessage', 'errorTitle', 'error',
                     'promptTitle', 'prompt'):
            value = getattr(self, attr)
            if value is not None:
                yield attr, safe_string(value)


class ValidationType(object):
    NONE = "none"
    WHOLE = "whole"
    DECIMAL = "decimal"
    LIST = "list"
    DATE = "date"
    TIME = "time"
    TEXT_LENGTH = "textLength"
    CUSTOM = "custom"


class ValidationOperator(object):
    BETWEEN = "between"
    NOT_BETWEEN = "notBetween"
    EQUAL = "equal"
    NOT_EQUAL = "notEqual"
    LESS_THAN = "lessThan"
    LESS_THAN_OR_EQUAL = "lessThanOrEqual"
    GREATER_THAN = "greaterThan"
    GREATER_THAN_OR_EQUAL = "greaterThanOrEqual"


class ValidationErrorStyle(object):
    STOP = "stop"
    WARNING = "warning"
    INFORMATION = "information"


def writer(data_validation):
    """
    Serialse a data validation
    """
    attrs = dict(data_validation)
    el = Element("dataValidation", attrs)
    for attr in ("formula1", "formula2"):
        value = getattr(data_validation, attr, None)
        if value is not None:
            f = Element(attr)
            f.text = value
            el.append(f)
    return el


def parser(element):
    """
    Parse dataValidation tag
    """
    dv = DataValidation(**element.attrib)
    for attr in ("formula1", "formula2"):
        for f in safe_iterator(element, attr):
            setattr(dv, attr, f.text)
    return dv
