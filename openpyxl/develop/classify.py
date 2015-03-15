from __future__ import absolute_import, print_function
# Copyright (c) 2010-2015 openpyxl

"""
Generate Python classes from XML Schema
"""

import argparse
import re

from openpyxl.tests.schema import (
    sheet_src,
    chart_src,
    drawing_main_src,
    shared_src,
    )

from lxml.etree import parse


XSD = "http://www.w3.org/2001/XMLSchema"

simple_mapping = {
    'xsd:boolean':'Bool',
    'xsd:unsignedInt':'Integer',
    'xsd:int':'Integer',
    'xsd:double':'Float',
    'xsd:string':'String',
    'xsd:unsignedByte':'Integer',
    'xsd:byte':'Integer',
    'xsd:long':'Float',
    'xsd:token':'String',
    's:ST_Panose':'HexBinary',
    's:ST_Lang':'String',
    'ST_Percentage':'String',
    'ST_PositivePercentage':'Percentage',
    'ST_TextPoint':'TextPoint',
    'ST_UniversalMeasure':'UniversalMeasure',
    'ST_Coordinate32':'Coordinate',
    'ST_Coordinate':'Coordinate',
    'ST_Coordinate32Unqualified':'Coordinate',
    's:ST_Xstring':'String',
    'ST_Angle':'Integer',
}

complex_mapping = {
    'Boolean':'Bool',
    'Double':'Float',
    'Long':'Integer',
}


ST_REGEX = re.compile("(?P<schema>[a-z]:)?(?P<typename>ST_[A-Za-z]+)")

def classify(tagname, src=sheet_src, schema=None):
    """
    Generate a Python-class based on the schema definition
    """
    if schema is None:
        schema = parse(src)
    nodes = schema.iterfind("{%s}complexType" % XSD)
    tag = None
    for node in nodes:
        if node.get('name') == tagname:
            tag = tagname
            break
    if tag is None:
        pass
        raise ValueError("Tag {0} not found".format(tagname))

    types = set()

    s = """\n\nclass %s(Serialisable):\n\n""" % tagname[3:]
    attrs = []

    # attributes
    for el in node.iterfind("{%s}attribute" % XSD):
        attr = el.attrib
        if 'ref' in attr:
            continue
        attrs.append(attr['name'])

        if attr.get("use") == "optional":
            attr["use"] = "allow_none=True"
        else:
            attr["use"] = ""
        if attr.get("type").startswith("ST_"):
            attr['type'] = simple(attr.get("type"), schema, attr['use'])
            types.add(attr['type'].split("(")[0])
            s += "    {name} = {type}\n".format(**attr)
        else:
            if attr['type'] in simple_mapping:
                attr['type'] = simple_mapping[attr['type']]
                types.add(attr['type'])
                s += "    {name} = {type}({use})\n".format(**attr)
            else:
                s += "    {name} = Typed(expected_type={type}, {use})\n".format(**attr)

    children = []
    elements = []
    child_elements = node.findall("{%s}sequence/{%s}element" % (XSD, XSD))
    for el in node.iterfind("{%s}sequence/{%s}element" % (XSD, XSD)):
        attr = {'name': el.get("name"),}

        typename = el.get("type")
        match = ST_REGEX.match(typename)
        if typename.startswith("xsd:"):
            attr['type'] = simple_mapping[typename]
            types.add(attr['type'])
        elif match is not None:
            src = srcs_mapping.get(match.group('schema'))
            if src is not None:
                schema = parse(src)
            typename = match.group('typename')
            attr['type'] = simple(typename, schema)
        else:
            children.append(typename)
            if (typename.startswith("a:")
                or typename.startswith("s:")
                ):
                attr['type'] = typename[5:]
            else:
                attr['type'] = typename[3:]

        attr['use'] = ""
        if el.get("minOccurs") == "0":
            attr['use'] = "allow_none=True"
        attrs.append(attr['name'])
        if attr['type'] in complex_mapping:
            attr['type'] = complex_mapping[attr['type']]
            s += "    {name} = {type}(nested=True, {use})\n".format(**attr)
        else:
            s += "    {name} = Typed(expected_type={type}, {use})\n".format(**attr)

    #if children:
        #names = (c.name for c in children)
        #s += "\n__elements__ = {0}\n".format(tuple(names))

    if attrs:
        s += "\n    def __init__(self,\n"
        for a in attrs:
            s += "                 %s=None,\n" % a
        s += "                ):\n"
    else:
        s += "    pass"
    for attr in attrs:
        s += "        self.{0} = {0}\n".format(attr)

    return s, types, children


def simple(tagname, schema, use=""):

    for node in schema.iterfind("{%s}simpleType" % XSD):
        if node.get("name") == tagname:
            break
    constraint = node.find("{%s}restriction" % XSD)
    if constraint is None:
        return "unknown defintion for {0}".format(tagname)
    typ = constraint.get("base")
    typ = "{0}()".format(simple_mapping.get(typ, typ))
    values = constraint.findall("{%s}enumeration" % XSD)
    values = [v.get('value') for v in values]
    if values:
        s = "Set"
        if "none" in values:
            idx = values.index("none")
            del values[idx]
            s = "NoneSet"
        typ = s + "(values=({0}))".format(values)
    return typ

srcs_mapping = {'a:':drawing_main_src, 's:':shared_src}

class ClassMaker:
    """
    Generate
    """

    def __init__(self, tagname, src=chart_src, classes=set()):
        self.schema=parse(src)
        self.types = set()
        self.classes = classes
        self.body = ""
        self.create(tagname)

    def create(self, tagname):
        body, types, children = classify(tagname, schema=self.schema)
        self.body = body + self.body
        self.types = self.types.union(types)
        for child in children:
            if (child.startswith("a:")
                or child.startswith("s:")
                ):
                src = srcs_mapping[child[:2]]
                tagname = child[2:]
                if tagname not in self.classes:
                    cm = ClassMaker(tagname, src=src, classes=self.classes)
                    self.body = cm.body + self.body # prepend dependent types
                    self.types.union(cm.types)
                    self.classes.add(tagname)
                    self.classes.union(cm.classes)
                continue
            if child not in self.classes:
                self.create(child)
                self.classes.add(child)

    def __str__(self):
        s = """#Autogenerated schema
        from openpyxl.descriptors.serialisable import Serialisable
        from openpyxl.descriptors import (\n    Typed,"""
        for t in self.types:
            s += "\n    {0},".format(t)
        s += (")\n")
        s += self.body
        return s


def make(element, schema=sheet_src):
    cm = ClassMaker(element, schema)
    print(cm)


commands = argparse.ArgumentParser(description="Generate Python classes for a specific scheme element")
commands.add_argument('element', help='The XML type to be converted')
commands.add_argument('--schema',
                      help='The relevant schema. The default is for worksheets',
                      choices=["sheet_src", "chart_src", "shared_src"],
                      default="sheet_src",
                      )

if __name__ == "__main__":
    args = commands.parse_args()
    schema = globals().get(args.schema)
    make(args.element, schema)
