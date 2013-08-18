def namedtuple(name, fields):
    return type(name, (object,), {'__slots__': fields})
