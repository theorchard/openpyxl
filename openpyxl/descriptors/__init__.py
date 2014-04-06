# Copyright (c) 2010-2014 openpyxl


"""Based on Python Cookbook 3rd Edition, 8.13"""


class Descriptor(object):

    def __init__(self, name=None, **kw):
        self.name = name
        for k, v in kw.items():
            setattr(self, k, v)

    def __set__(self, instance, value):
        instance.__dict__[self.name] = value


class Typed(Descriptor):

    expected_type = type(None)

    def __set__(self, instance, value):
        if not isinstance(value, self.expected_type):
            try:
                value = self.expected_type(value)
            except:
                raise TypeError('expected ' + str(self.expected_type))
        super(Typed, self).__set__(instance, value)


class Max(Descriptor):

    def __init__(self, name=None, **kw):
        if 'max' not in kw:
            raise TypeError('missing max value')
        super(Max, self).__init__(name, **kw)

    def __set__(self, instance, value):
        if value > self.max:
            raise ValueError('Max value is {0}'.format(self.max))
        super(Max, self).__set__(instance, value)


class Min(Descriptor):

    def __init__(self, name=None, **kw):
        if 'min' not in kw:
            raise TypeError('missing min value')
        super(Min, self).__init__(name, **kw)

    def __set__(self, instance, value):
        if value < self.min:
            raise ValueError('Min value is {0}'.format(self.min))
        super(Min, self).__set__(instance, value)


class MinMax(Min, Max):
    pass


class Integer(Typed):

    expected_type = int


class Float(Typed):

    expected_type = float


class Bool(Typed):

    expected_type = bool


class Set(Descriptor):
    """Value can only be from a set of know values"""

    def __init__(self, name=None, **kw):
        if not 'values' in kw:
            raise TypeError("missing set of values")
        super(Set, self).__init__(name, **kw)

    def __set__(self, instance, value):
        if value not in self.values:
            raise ValueError("Value must be one of {0}".format(self.values))
        super(Set, self).__set__(instance, value)


class MetaStrict(type):

    def __new__(cls, clsname, bases, methods):
        for k, v in methods.items():
            if isinstance(v, Descriptor):
                v.name = k
        return type.__new__(cls, clsname, bases, methods)

Strict = MetaStrict('Strict', (object, ), {}
               )

del MetaStrict
