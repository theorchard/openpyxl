from __future__ import absolute_import
# Copyright (c) 2010-2014 openpyxl


import os
from zipfile import ZipFile

from lxml.etree import XMLSchema
from lxml.etree import parse

# Provide schema based validators, lxml required
# use schema.validate(Element) or schema.assertValid(Element) for messages

SCHEMA_FOLDER = os.path.join(os.path.dirname(__file__), 'schemas')

sheet_src = os.path.join(SCHEMA_FOLDER, 'sml.xsd')
sheet_schema = XMLSchema(file=sheet_src)

chart_src = os.path.join(SCHEMA_FOLDER, 'dml-chart.xsd')
chart_schema = XMLSchema(file=chart_src)

drawing_src = os.path.join(SCHEMA_FOLDER, 'dml-spreadsheetDrawing.xsd')
drawing_schema = XMLSchema(file=drawing_src)

drawing_main_src = os.path.join(SCHEMA_FOLDER, "dml-main.xsd")

core_src = os.path.join(SCHEMA_FOLDER, 'opc-coreProperties.xsd')
core_props_schema = XMLSchema(file=core_src)

sml_files = ['xl/styles.xml']  # , 'xl/workbook.xml']


def validate_archive(file_path):
    zipfile = ZipFile(file_path)
    try:
        for entry in zipfile.infolist():
            filename = entry.filename
            f = zipfile.open(entry)
            root = parse(f).getroot()
            if filename in sml_files or filename.startswith('xl/worksheets/sheet'):
                if root.get('{http://www.w3.org/XML/1998/namespace}space'):
                    # not allowed by schema
                    del root.attrib['{http://www.w3.org/XML/1998/namespace}space']
                sheet_schema.assertValid(root)
    finally:
        zipfile.close()


XSD = "http://www.w3.org/2001/XMLSchema"

mapping = {
    'xsd:boolean':'Bool',
    'xsd:unsignedInt':'Integer',
    'xsd:int':'Integer',
    'xsd:double':'Float',
    'xsd:string':'String',
    'xsd:unsignedByte':'MinMax',
    'xsd:byte':'MinMax',
    'xsd:long':'Integer',
    'xsd:token':'String',
}

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

    s = """\n\nclass %s(Strict):\n\n""" % tagname[3:]
    attrs = []

    # attributes
    for el in node.iterfind("{%s}attribute" % XSD):
        attr = el.attrib
        if 'ref' in attr:
            continue
        attrs.append(attr['name'])
        if attr['type'] in mapping:
            attr['type'] = mapping[attr['type']]
            types.add(attr['type'])
        if attr.get("use") == "optional":
            attr["use"] = "allow_none=True"
        else:
            attr["use"] = ""
        if attr.get("type").startswith("ST_"):
            attr['type'] = simple(attr.get("type"), schema)
            types.add(attr['type'].split("(")[0])
            s += "    {name} = {type}\n".format(**attr)
        else:
            s += "    {name} = {type}({use})\n".format(**attr)

    children = []
    for el in node.iterfind("{%s}sequence/{%s}element" % (XSD, XSD)):
        attr = {'name': el.get("name"),}

        typename = el.get("type")
        if typename.startswith("xsd:"):
            attr['type'] = mapping[typename]
            types.add(attr['type'])
        else:
            children.append(typename)
            if typename.startswith("a:"):
                attr['type'] = typename[5:]
            else:
                attr['type'] = typename[3:]

        attr['use'] = ""
        if el.get("minOccurs") == "0":
            attr['use'] = "allow_none=True"
        attrs.append(attr['name'])
        s += "    {name} = {type}({use})\n".format(**attr)

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


def simple(tagname, schema):

    for node in schema.iterfind("{%s}simpleType" % XSD):
        if node.get("name") == tagname:
            break
    constraint = node.find("{%s}restriction" % XSD)
    if constraint is None:
        return "unknown defintion for {0}".format(tagname)
    typ = constraint.get("base")
    typ = "{0}()".format(mapping.get(typ, typ))
    values = constraint.findall("{%s}enumeration" % XSD)
    values = [v.get('value') for v in values]
    if values:
        typ = "Set(values=({0}))".format(values)
    return typ
