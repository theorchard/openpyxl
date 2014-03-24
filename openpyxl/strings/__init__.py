# Copyright (c) 2010-2014 openpyxl


class IndexedList(list):
    """
    List with optimised access by value
    Based on Alex Martelli's recipe

    http://code.activestate.com/recipes/52303-the-auxiliary-dictionary-idiom-for-sequences-with-/
    """

    def __init__(self, iterable=()):
        self.clean = False
        if iterable != ():
            self._dict = dict.fromkeys(iterable)
            iterable = sorted(self._dict)
            self.clean = True
        super(IndexedList, self).__init__(iterable)

    def _rebuild_dict(self):
        self._dict = {}
        idx = 0
        for value in self:
            if value not in self._dict:
                self._dict[value] = idx
                idx += 1

    def __contains__(self, value):
        if not self.clean:
            self._rebuild_dict()
        return value in self._dict

    def index(self, value):
        if value in self:
            return self._dict[value]
        raise ValueError

    @property
    def values(self):
        """Return a deduped sorted list of the list's values"""
        if not self.clean:
            self._rebuild_dict()
        return sorted(self._dict)


def method_wrapper(methodname):
    _method = getattr(list, methodname)
    def wrapper(self, *args):
        self.clean = False
        return _method(self, *args)
    setattr(IndexedList, methodname, wrapper)

for meth in ['__setitem__', '__delitem__', '__setslice__', '__delslice__',
             '__iadd__', 'insert', 'append', 'pop', 'remove', 'extend', 'sort']:
    method_wrapper(meth)

del method_wrapper
