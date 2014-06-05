from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl


from .. password_hasher import hash_password


def test_password():
    enc = hash_password('secret')
    assert enc == 'DAA7'
