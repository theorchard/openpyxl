from openpyxl import LXML
from openpyxl.xml.functions import register_namespace
from openpyxl.xml.constants import SHEET_MAIN_NS, REL_NS

def lxml_fix():
    from lxml.etree import parse
    tree = parse("Issues/workbook-broken-in-ios.xml")
    tree.write("Issues/workbook-fixed.xml")


def etree_fix():
    from xml.etree.ElementTree import parse
    tree = parse("Issues/workbook-broken-in-ios.xml")
    tree.write("Issues/workbook-fixed.xml")


if __name__ == "__main__":
    from xml.etree.ElementTree import register_namespace, parse
    register_namespace("", SHEET_MAIN_NS)
    register_namespace("r", REL_NS)
    tree = parse("Issues/workbook-broken-in-ios.xml")
    tree.write("Issues/workbook-fixed.xml")
