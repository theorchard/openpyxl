# Copyright (c) 2010-2014 openpyxl

import pytest

@pytest.fixture
def Alignment():
    from .. alignment import Alignment
    return Alignment


def test_default(Alignment):
    al = Alignment()
    assert dict(al) == {'horizontal':'general', 'vertical':'bottom'}


def test_round_trip(Alignment):
    args = {'horizontal':'center', 'vertical':'top', 'textRotation':'45', 'indent':'4'}
    al = Alignment(**args)
    assert dict(al) == args


def test_alias(Alignment):
    al = Alignment(text_rotation=90, shrink_to_fit=True, wrap_text=True)
    assert dict(al) == {'horizontal':'general', 'vertical':'bottom',
                        'textRotation':'90',
                        'shrinkToFit':'1',
                        'wrapText':'1'}
