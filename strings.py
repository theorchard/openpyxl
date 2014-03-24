import random
from openpyxl.writer.strings import StringTableBuilder

from openpyxl.strings import IndexedList

chars = [chr(random.randint(32, 255)) for i in range(1000)]
chars.sort()

def check1():
    table = StringTableBuilder()
    for c in chars:
        table.add(c)
    return sorted(table.get_table())

def check2():
    table = IndexedList((c.strip() for c in chars))
    return table.values


assert check1() == check2()
