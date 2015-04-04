from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from io import BytesIO
from tempfile import NamedTemporaryFile
from zipfile import BadZipfile

from openpyxl.utils.exceptions import InvalidFileException
from .. excel import load_workbook

import pytest


def test_read_empty_file(datadir):
    datadir.chdir()
    with pytest.raises(BadZipfile):
        load_workbook('null_file.xlsx')

def test_load_workbook_from_fileobj(datadir):
    """ can a workbook be loaded from a file object without exceptions
    This tests for regressions of
    https://bitbucket.org/openpyxl/openpyxl/issue/433
    """
    datadir.chdir()
    with open('empty_with_no_properties.xlsx', 'rb') as f:
        load_workbook(f)

def test_repair_central_directory():
    from ..excel import repair_central_directory, CENTRAL_DIRECTORY_SIGNATURE

    data_a = b"foobarbaz" + CENTRAL_DIRECTORY_SIGNATURE
    data_b = b"bazbarfoo1234567890123456890"

    # The repair_central_directory looks for a magic set of bytes
    # (CENTRAL_DIRECTORY_SIGNATURE) and strips off everything 18 bytes past the sequence
    f = repair_central_directory(BytesIO(data_a + data_b), True)
    assert f.read() == data_a + data_b[:18]

    f = repair_central_directory(BytesIO(data_b), True)
    assert f.read() == data_b


@pytest.mark.parametrize("extension",
                         ['.xlsb', '.xls', 'no-format']
                         )
def test_invalid_file_extension(extension):
    tmp = NamedTemporaryFile(suffix=extension)
    with pytest.raises(InvalidFileException):
        load_workbook(filename=tmp.name)


def test_style_assignment(datadir):
    from ..excel import load_workbook

    datadir.chdir()
    wb = load_workbook("complex-styles.xlsx")
    assert len(wb._alignments) == 9
    assert len(wb._fills) == 6
    assert len(wb._fonts) == 8
    assert len(wb._borders) == 7
    assert len(wb._number_formats) == 0
    assert len(wb._protections) == 1


def test_read_stringio():
    filelike = BytesIO(b"certainly not a valid XSLX content")
    # Test invalid file-like objects are detected and not handled as regular files
    with pytest.raises(BadZipfile):
        load_workbook(filelike)
