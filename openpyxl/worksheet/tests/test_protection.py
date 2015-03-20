from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

import pytest


from .. protection import SheetProtection, hash_password


def test_password():
    enc = hash_password('secret')
    assert enc == 'DAA7'


def test_ctor():
    prot = SheetProtection()
    assert dict(prot) == {
        'autoFilter': '1', 'deleteColumns': '1',
        'deleteRows': '1', 'formatCells': '1', 'formatColumns': '1', 'formatRows':
        '1', 'insertColumns': '1', 'insertHyperlinks': '1', 'insertRows': '1',
        'objects': '0', 'pivotTables': '1', 'scenarios': '0', 'selectLockedCells':
        '0', 'selectUnlockedCells': '0', 'sheet': '0', 'sort': '1'
    }


def test_ctor_with_password():
    prot = SheetProtection(password="secret")
    assert prot.password == "DAA7"


@pytest.mark.parametrize("password, already_hashed, value",
                         [
                             ('secret', False, 'DAA7'),
                             ('secret', True, 'secret'),
                         ])
def test_explicit_password(password, already_hashed, value):
    prot = SheetProtection()
    prot.set_password(password, already_hashed)
    assert prot.password == value
    assert prot.sheet == True
