try:
    from itertools import ifilter
except ImportError:
    ifilter = filter

try:
    from itertools import izip
except ImportError:
    izip = zip

try:
    range = xrange
except NameError:
    range = range

def iteritems(iterable):
    if hasattr(iterable, 'iteritems'):
        for item in iterable.iteritems():
            yield item
    else:
        for item in iterable.items():
            yield item

def iterkeys(iterable):
    if hasattr(iterable, 'iterkeys'):
        for item in iterable.iterkeys():
            yield item
    else:
        for item in iterable.keys():
            yield item


def itervalues(iterable):
    if hasattr(iterable, 'itervalues'):
        for item in iterable.itervalues():
            yield item
    else:
        for item in iterable.values():
            yield item
