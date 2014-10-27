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
}

def classify(tagname):
    """
    Generate a Python-class based on the schema definition
    """
    schema = parse(sheet_src)
    nodes = schema.iterfind("{%s}complexType" % XSD)
    for node in nodes:
        if node.get('name') == tagname:
            break

    types = set()

    s = """
from openpyxl.descriptors import Strict


class %s(Strict):
""" % tagname[3:]
    attrs = []

    for el in node.iterfind("{%s}attribute" % XSD):
        attr = el.attrib
        if 'ref' in attr:
            continue
        attrs.append(attr['name'])
        if attr['type'] in mapping:
            attr['type'] = mapping[attr['type']]
            types.add(mapping[attr['type']])
        if attr.get("use") == "optional":
            attr["use"] = "allow_none=True"
        else:
            attr["use"] = ""
        if attr.get("type").startswith("ST_"):
            s += "    " + simple(attr.get("type"))
        s += "    {name} = {type}({use})\n".format(**attr)

    s += "\n"

    for el in node.iterfind("{%s}sequence/{%s}element" % (XSD, XSD)):
        s += "    {name} = {type}()\n".format(name=el.get("name"), type=el.get("type")[3:])

    s += "    def __init__(self,\n    %s=None):\n" % ("=None,\n    ".join(attrs))
    for attr in attrs:
        s += "    {0} = {0}\n".format(attr)

    for el in node.iterfind("{%s}sequence/{%s}element" % (XSD, XSD)):
        s += "\n\n"
        s += classify(el.get('type'))
    return s


def simple(tagname):

    schema = parse(sheet_src)
    for node in schema.iterfind("{%s}simpleType" % XSD):
        if node.get("name") == tagname:
            break
    constraint = node.find("{%s}restriction" % XSD)
    typ = constraint.get("base")
    typ = "{0}()".format(mapping.get(typ, typ))
    values = constraint.findall("{%s}enumeration" % XSD)
    values = [v.get('value') for v in values]
    if values:
        typ = "Set(values=({0}))".format(values)
    return "{0} = {1}\n".format(tagname[3:], typ)
