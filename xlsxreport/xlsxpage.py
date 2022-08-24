class XlsxPage:
    def __init__(self, page_name):
        self.page_name = page_name
        self.tables = []
        self.titles = []
        self.titles_at(0, 0)
        self._workbook = None
        self._worksheet = None


    def add_page_table(self, name, table):
        self.tables.append(
            {'name': name})
        setattr(self, name, table)


    def set_workbook_worksheet(self, wb, ws):
        self._workbook = wb
        self._worksheet = ws


    def titles_at(self, row:int=0, col:int=0):
        """
        Set top, left cell for first page title (next on new row)
        """
        self._titles_at_row = row
        self._titles_at_col = col


    def write_page(self):
        if not self.titles:
            return
        row = self._titles_at_row
        for title in self.titles:
            self._worksheet.write(
                row, self._titles_at_col, title)
            row += 1
