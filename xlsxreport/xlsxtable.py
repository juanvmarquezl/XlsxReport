import uuid

class XlsxTable:
    """
    #Declares an excel table

    ##cols_setup:
    This dict contains all cell (column) setup

        table.cols_setup = {
            'col_dict_key': {
                'width': 30,  # Column with (xls with units)
                'type': float,  # Python cell's data type
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
            },
            ...
        }

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


    def _set_workbook_formats(self):
        self.formats = {}
        for key in self.cols_setup.keys():
            format = self.cols_setup[key].get('format', None)
            self.formats.update(
                {key: self._workbook.add_format(format)})
        return True


    def _get_format(self, key):
        return self.formats.get(key)


    def _write_headers(self, row):
        self.header_row = row
        if not self.headers and self.table_data:
            self.headers = self.table_data[0].keys()
        for col in range(len(self.headers)):
            self._worksheet.write(
                row, self.start_at_col + col, list(self.headers)[col])


    def _convert_cell_value(self, value, type):
        return type(value) if type else value


    def _write_data_row(self, row, line):
        if self.first_row == None:
            self.first_row = row
        col = self.start_at_col

        for key, val in line.items():
            type = self.cols_setup.get(key, {}).get('type')
            cell_val = self._convert_cell_value(val, type)
            self._worksheet.write(
                row, col, cell_val, self._get_format(key))
            col +=1
        self.last_row = row


    def set_workbook_worksheet(self, wb, ws):
        self._workbook = wb
        self._worksheet = ws


    def before_write_table(self):
        if not self.cols_setup and self.table_data:
            self.cols_setup = {k: {} for k in self.table_data[0].keys()}
        col = self.start_at_col
        # Set columns width to default
        for key, value in self.cols_setup.items():
            if value.get('width'):
                self._worksheet.set_column(col, col, value['width'])
            col += 1
        return True


    def write_table(self):
        row = self.start_at_row
        self._set_workbook_formats()
        for item in self.table_data:
            if row == self.start_at_row:
                self._write_headers(row)
            row += 1
            self._write_data_row(row, item)

