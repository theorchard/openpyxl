# Copyright (c) 2010-2014 openpyxl

import py.test


from .. protection import SheetProtection


def test_ctor():
    prot = SheetProtection()
    assert dict(prot) == {
        'autoFilter': '1', 'deleteColumns': '1',
        'deleteRows': '1', 'formatCells': '1', 'formatColumns': '1', 'formatRows':
        '1', 'insertColumns': '1', 'insertHyperlinks': '1', 'insertRows': '1',
        'objects': '0', 'pivotTables': '1', 'scenarios': '0', 'selectLockedCells':
        '0', 'selectUnlockedCells': '0', 'sheet': '0', 'sort': '1'
    }

