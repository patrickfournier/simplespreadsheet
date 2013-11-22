# -*- coding: utf-8 -*-
"""
Register a simplespreadsheet directive to do simple calculations in a table.

Create a table with the sum, the difference, the product and the
quotient of 2 and 3 in the third column:

  .simplespreadsheet::
      +------+------+-------------+
      | ={2} | ={3} | ={a# + b#}  |
      +------+------+-------------+
      | ={2} | ={3} | ={a# - b#}  |
      +------+------+-------------+
      | ={2} | ={3} | ={a# * b#}  |
      +------+------+-------------+
      | ={2} | ={3} | ={a# / b#}  |
      +------+------+-------------+

See SpreadsheetDirective.resolve documentation for more information on
formula syntax.

For now, the only available function is sum("a1:c6"). You can easily
add new functions by adding them to SpreadSheet.tools
"""

import re
import docutils
import rst2pdf


class SpreadSheet:
    """
    Simple spreadsheet from
    http://code.activestate.com/recipes/355045-spreadsheet/
    """
    _cells = {}
    tools = {}

    def __init__(self):
        SpreadSheet.tools['sum'] = self.sum

    def __setitem__(self, key, formula):
        self._cells[key] = formula

    def getformula(self, key):
        return self._cells[key]

    def __getitem__(self, key ):
        return eval(self._cells[key], SpreadSheet.tools, self)

    @staticmethod
    def col_coords(col):
        # Compute alphabetical representation of col.
        s = ''
        while col >= 0:
            digit = col % 26
            s = chr(ord('a') + digit) + s
            col = col / 26 - 1

        return s

    @staticmethod
    def row_coords(row):
        return str(row+1)

    @staticmethod
    def coords(col, row):
        """
        Utility function to return spreadsheet index (1-based) from
        row and column index (zero-based).
        """
        return SpreadSheet.col_coords(col) + SpreadSheet.row_coords(row)

    @staticmethod
    def inv_coords(c):
        """
        Utility function to return row and column index (0-based) from
        spreadsheet index (1-based).
        """
        r = re.search('([a-z]+)([0-9]+)', c);
        col_str = r.group(1)
        row = int(r.group(2)) - 1

        col = 0
        m = 1
        while col_str != '':
            col = col + (ord(col_str[-1]) - ord('a') + 1) * m
            col_str = col_str[:-1]
            m = m * 26
        col = col - 1

        return (col, row)

    def sum(self, cells):
        """
        Return the sum of the cells in the range given by *cells*.

        Example:
          Spreadsheet.sum('a1:c5')

        Range limits are included in sum.
        """
        (first, last) = cells.split(':')

        (c1, l1) = SpreadSheet.inv_coords(first)
        (c2, l2) = SpreadSheet.inv_coords(last)

        s = 0
        for c in xrange(c1, c2+1):
            for l in xrange(l1, l2+1):
                ax = SpreadSheet.coords(c, l)
                s += self[ax]
        return s;


class SpreadsheetNode(docutils.nodes.General, docutils.nodes.Element):
    pass


class SpreadsheetDirective(docutils.parsers.rst.Directive):
    """Directive to insert spreadsheet markup."""
    required_arguments = 0
    optional_arguments = 0
    option_spec = {}
    has_content = True

    def run(self):
        self.assert_has_content()
        node = SpreadsheetNode()
        self.state.nested_parse(self.content, self.content_offset, node)
        self.resolve(node)
        return [node]

    def resolve(self, ssn):
        """
        Resolve formulae in SpreadsheetNode ssn, replacing them by a
        computed value. The SpreadsheetNode must contains a table node
        as its first child.

        Formulae are piece of text enclosed in ={...}. Oly one formula
        per cell is supported.

        A simple formula is simply a number, like this:

          ={4}

        Some more complex formulae:

          ={a1 * a2}
          ={b# * f#}
          ={@1 + @2}
          ={sum("a1:a20")}

        Special characters @ and # are replaced by current column and row,
        respectively.

        A cell has value 0 (zero) unless it contains a formula. In
        this case, the cell takes the value of the formula's result.
        """

        # Find table body.
        # ssn[0] is the table node
        tgroup_idx = ssn[0].first_child_matching_class(docutils.nodes.tgroup)
        tbody_idx = ssn[0][tgroup_idx].first_child_matching_class(docutils.nodes.tbody)
        body = ssn[0][tgroup_idx][tbody_idx]

        s = SpreadSheet()

        # Copy values and formulae from table to spreadsheet.
        for row in xrange(len(body)):
            offset = 0
            if type(body[row]) == docutils.nodes.row:
                for col in xrange(len(body[row])):
                    if type(body[row][col]) == docutils.nodes.entry:
                        ax = s.coords(col + offset, row)
                        text = self.parse_entry(body[row][col], '0')
                        text = text.replace('@', s.col_coords(col + offset))
                        text = text.replace('#', s.row_coords(row))
                        s[ax] = text

                        # Increment column address if cell spans more
                        # than one column.
                        if body[row][col].hasattr('morecols'):
                            offset += body[row][col].get('morecols')

        # Copy values from spreadsheet to table.
        for row in xrange(len(body)):
            offset = 0
            if type(body[row]) == docutils.nodes.row:
                for col in xrange(len(body[row])):
                    if type(body[row][col]) == docutils.nodes.entry:
                        ax = s.coords(col + offset, row)
                        self.replace_value(ssn[0][tgroup_idx][tbody_idx][row][col], s[ax])

                        # Increment column address if cell spans more
                        # than one column.
                        if body[row][col].hasattr('morecols'):
                            offset += body[row][col].get('morecols')


    def parse_entry(self, entry, default_text):
        """
        Extract formula from table cell *entry*.

        If the cell does not contain any formula, return
        *default_text*.
        """
        for i in xrange(len(entry)):
            if type(entry[i]) == docutils.nodes.Text:
                r = re.search('={(.*)}', entry[i].astext())
                if r:
                    default_text = r.group(1)
            else:
                default_text = self.parse_entry(entry[i], default_text)

        return default_text


    def replace_value(self, entry, value):
        """Replace formulae with their value."""
        for i in xrange(len(entry)):
            if type(entry[i]) == docutils.nodes.Text:
                (new_text, subs_count) = re.subn('={.*}', str(value), entry[i].astext())
                if subs_count > 0:
                    new_node = docutils.nodes.Text(new_text)
                    entry.replace(entry[i], new_node)
            else:
                self.replace_value(entry[i], value)


class SpreadsheetHandler(rst2pdf.genelements.NodeHandler, SpreadsheetNode):
    pass


docutils.parsers.rst.directives.register_directive("simplespreadsheet", SpreadsheetDirective)
