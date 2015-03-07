from __future__ import absolute_import
# Copyright (c) 2010-2015 openpyxl

from openpyxl.compat import iteritems, OrderedDict, deprecated

from openpyxl.styles import Font, PatternFill, Border
from openpyxl.styles.differential import DifferentialFormat
from .rules import CellIsRule, ColorScaleRule, FormatRule, FormulaRule


def unpack_rules(cfRules):
    for key, rules in iteritems(cfRules):
        for idx,rule in enumerate(rules):
            yield (key, idx, rule['priority'])


class ConditionalFormatting(object):
    """Conditional formatting rules."""
    rule_attributes = ('aboveAverage', 'bottom', 'dxfId', 'equalAverage',
                       'operator', 'percent', 'priority', 'rank', 'stdDev', 'stopIfTrue',
                       'text')
    icon_attributes = ('iconSet', 'showValue', 'reverse')

    def __init__(self):
        self.cf_rules = OrderedDict()
        self.max_priority = 0
        self.parse_rules = {}

    def add(self, range_string, cfRule):
        """Add a rule.  Rule is either:
         1. A dictionary containing a key called type, and other keys, as in `ConditionalFormatting.rule_attributes`.
         2. A rule object, such as ColorScaleRule, FormulaRule or CellIsRule

         The priority will be added automatically.
        """
        if isinstance(cfRule, dict):
            rule = cfRule
        else:
            rule = cfRule.rule
        rule['priority'] = self.max_priority + 1
        self.max_priority += 1
        if range_string not in self.cf_rules:
            self.cf_rules[range_string] = []
        self.cf_rules[range_string].append(rule)


    def _fix_priorities(self):
        rules = unpack_rules(self.cf_rules)
        rules = sorted(rules, key=lambda x: x[2])
        for idx, (key, rule_no, prio) in enumerate(rules, 1):
            self.cf_rules[key][rule_no]['priority'] = idx
        self.max_priority = len(rules)


    def update(self, cfRules):
        """Set the conditional formatting rules from a dictionary.  Intended for use when loading a document.
        cfRules use the structure: {range_string: [rule1, rule2]}, eg:
        {'A1:A4': [{'type': 'colorScale', 'priority': 13, 'colorScale': {'cfvo': [{'type': 'min'}, {'type': 'max'}],
        'color': [Color('FFFF7128'), Color('FFFFEF9C')]}]}
        """
        for range_string, rules in iteritems(cfRules):
            if range_string not in self.cf_rules:
                self.cf_rules[range_string] = rules
            else:
                self.cf_rules[range_string] += rules
        self._fix_priorities()


    @deprecated("Conditionl Formats are saved automatically")
    def setDxfStyles(self, wb):
        pass
