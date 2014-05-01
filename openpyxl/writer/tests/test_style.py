# Copyright (c) 2010-2014 openpyxl

import pytest

from openpyxl.styles.fills import GradientFill
from openpyxl.styles.colors import Color
from openpyxl.writer.styles import StyleWriter
from openpyxl.tests.helper import get_xml, compare_xml


class DummyWorkbook:

    style_properties = []


def test_write_gradient_fill():
    fill = GradientFill(degree=90, stop=[Color('theme:0:'), Color('theme:4:')])
    writer = StyleWriter(DummyWorkbook())
    writer._write_gradient_fill(writer._root, fill)
    xml = get_xml(writer._root)
    expected = """<?xml version="1.0" ?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <gradientFill bottom="0" degree="90" left="0" right="0" top="0" type="linear">
    <color theme="0"/>
    <color theme="4"/>
  </gradientFill>
</styleSheet>
    """
    diff = compare_xml(xml, expected)
    assert diff is None, diff
