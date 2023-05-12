import uuid
import xlsxwriter
import re
from datetime import datetime




class XlsxTable:
    """
    This module provides a class for creating and writing Excel tables.

    The XlsxTable class represents an Excel table. It can be used to create and write Excel tables.

    Attributes

        table_data: The data for the table. This is a list of dictionaries, where each dictionary represents a row in the table.
        cols_setup: The column setup for the table. This is a dictionary, where the keys are the column names and the values are the column settings.
        _workbook: The Excel workbook object.
        _worksheet: The Excel worksheet object.
        start_at_row: The row number where the table should start.
        start_at_col: The column number where the table should start.
        headers: The headers for the table. This is a list of strings, where each string is the header for a column.
        header_row: The row number where the headers are located.
        first_row: The row number where the data starts.
        last_row: The row number where the data ends.

    Methods

        __init__(self, table_data): The constructor for the XlsxTable class.
        _set_workbook_formats(): Sets the workbook formats for the table.
        _get_format(self, key): Gets the format for a column.
        _write_headers(self, row): Writes the headers for the table.
        _convert_cell_value(self, value, type): Converts a cell value to the correct type.
        _gen_cell_formula(self, cell_value, cell_row): Generates a formula for a cell.
        _write_data_row(self, row, line): Writes a row of data to the table.
        set_workbook_worksheet(self, wb, ws): Sets the workbook and worksheet objects for the table.
        add_workbook_format(self, name, format): Adds a workbook format to the table.
        set_table_headers_format(self, name): Sets the format for the table headers.
        before_write_table(self): Performs actions before writing the table.
        write_table(self): Writes the table to the workbook.

    ##cols_setup:
    This dict contains all cell (column) setup

        table.cols_setup = {
            'col_dict_key': {
                'width': 30,  # Column with (xls with units)
                'type': float,  # Python cell's data type
                'title': str,  # Col's title
                'format': {
                    'num_format': '@',
                    'font_size': 10,
                    'font_color': 'gray',
                    'bg_color': '#FFEB9C'
                    'align': 'center',
                    'valign': 'vcenter',
                    'bold': True,
                    'text_wrap': True,
                    },  # cell format*
                'formula': '{{col_dict_key}}operator{{col_dict_key}}',
                'internal_link_c2p': create a internal link from cell to page
            },
            ...
        }

    formula:
        You can add column's formulas unsing double brackets {{col_dict_key}} to identify
        column with any valid excel formula, XlsxReport will change column identifier
        with excel's cell identifier. See /examples

    * for more info about cell format see:
        https://xlsxwriter.readthedocs.io/format.html
    """

    def __init__(self, table_data):
        self.id = str(uuid.uuid4())[:8]
        self.table_data = table_data
        self._workbook = None
        self._worksheet = None
        self.start_at_row = 0
        self.start_at_col = 0
        self.headers = []
        self.header_row = None
        self.first_row = None
        self.last_row = None
        self.cols_setup = {}
        self._formats = {}
        self._extra_formats = {}
        self._table_headers_format = None


    def _set_workbook_formats(self):
        '''
        '''
        for key in self.cols_setup.keys():
            format = self.cols_setup[key].get('format', None)
            self._formats.update(
                {key: self._workbook.add_format(format)})
        for key, format in self._extra_formats.items():
            self._formats.update(
                {key: self._workbook.add_format(format)})
        return True


    def _get_format(self, key):
        return self._formats.get(key)


    def _write_headers(self, row):
        """
        Write & format table headers
        """
        self.header_row = row
        format = self._get_format(self._table_headers_format)
        for col in range(len(self.headers)):
            self._worksheet.write(
                row, self.start_at_col + col, list(self.headers)[col], format)


    def _convert_cell_value(self, value, type):
        if type == datetime:
            return value
        return type(value) if type else value


    def _gen_cell_formula(self, cell_value, cell_row):
        '''
        Generate formula by replace {{col_dict_key}} with excel's cell ref

        General format:
            {{col_dict_key}}operator{{col_dict_key}}

        Samples:
            'formula': '={{col_dict_key1}}*{{col_dict_key2}}',
            'formula': '=SUM({{col_dict_key1}}:{{col_dict_key2}})',

        '''
        cell_formula = cell_value['formula']
        regex = '{{(.+?)}}'
        cols = re.findall(regex, cell_formula)
        for c in cols:
            repl_a = '{{%s}}' % c
            repl_b = self.cols_setup.get(c, {}).get('col_letter')
            if repl_a and repl_b:
                repl_b = repl_b + str(cell_row + 1)
            cell_formula = cell_formula.replace(repl_a, repl_b)
        return cell_formula if cell_formula[0] == '=' else '=' + cell_formula


    def _write_data_row(self, row, line):
        if self.first_row == None:
            self.first_row = row
        col = self.start_at_col

        for key, value in self.cols_setup.items():
            val  = line.get(key)
            type = value.get('type')
            if not value.get('formula'):  # get value
                if val:
                    cell_val = self._convert_cell_value(val, type)
                    if self.cols_setup[key].get('internal_link_c2p'):
                        _link = self.cols_setup[key].get('internal_link_c2p')
                        self._worksheet.write(
                            row, col, f"internal:'{cell_val}'!{_link}", self._get_format(key))
                    self._worksheet.write(
                        row, col, cell_val, self._get_format(key))
                else:
                    self._worksheet.write(
                        row, col, None, self._get_format(key))
            elif value.get('formula'):  # Write formula
                cell_formula = self._gen_cell_formula(value, row)
                self._worksheet.write(
                    row, col, cell_formula, self._get_format(key))
            col +=1
        self.last_row = row


    def set_workbook_worksheet(self, wb, ws):
        self._workbook = wb
        self._worksheet = ws


    def add_workbook_format(self, name, format):
        """
        Define a new format for cells
        """
        self._extra_formats.update(
            {name: format})
        return True


    def set_table_headers_format(self, name):
        """
        Assign format's name for table headers
        """
        self._table_headers_format = name
        return True


    def before_write_table(self):
        if not self.cols_setup and self.table_data:
            self.cols_setup = {k: {} for k in self.table_data[0].keys()}
        # Set table headers
        if not self.headers and self.table_data:
            self.headers = list(self.cols_setup.keys())
        # Set columns width to default
        col = self.start_at_col
        idx_col = 0
        for key, value in self.cols_setup.items():
            if value.get('width'):
                self._worksheet.set_column(col, col, value['width'])
            if value.get('title'):
                self.headers[idx_col] = value.get('title')
            value['col_letter'] = xlsxwriter.utility.xl_col_to_name(col)
            col += 1
            idx_col += 1

        return True


    def write_table(self):
        row = self.start_at_row
        self._set_workbook_formats()
        for item in self.table_data:
            if row == self.start_at_row:
                self._write_headers(row)
            row += 1
            self._write_data_row(row, item)

