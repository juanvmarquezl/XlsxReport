class XlsxPage:
    """
    This class is used to create and write Excel pages.

    Args:
        page_name (str): The name of the Excel page to create.

    Attributes:
        page_name (str): The name of the Excel page to create.
        tables (list): A list of tables to add to the Excel page.
        titles (list): A list of titles for the Excel page.
        titles_at_row (int): The row number where the titles should be written.
        titles_at_col (int): The column number where the titles should be written.
        _workbook (xlsxwriter.Workbook): The Excel workbook object.
        _worksheet (xlsxwriter.Worksheet): The Excel worksheet object.

    Methods:
        add_page_table(name, table): Add a table to the Excel page.
        set_workbook_worksheet(wb, ws): Set the Excel workbook and worksheet objects.
        titles_at(row, col): Set the row and column numbers where the titles should be written.
        after_page_titles(plus_rows): Get the row number after the titles.
        write_page(): Write the Excel page.
    """

    def __init__(self, page_name):
        """
        Create a new XlsxPage object.

        Args:
            page_name (str): The name of the Excel page to create.
        """
        self.page_name = page_name
        self.tables = []
        self.titles = []
        self.titles_at(0, 0)
        self._workbook = None
        self._worksheet = None


    def add_page_table(self, name, table):
        """
        Add a table to the Excel page.

        Args:
            name (str): The name of the table to add.
            table (object): The table object to add.
        """
        self.tables.append(
            {'name': name})
        setattr(self, name, table)


    def set_workbook_worksheet(self, wb, ws):
        """
        Set the Excel workbook and worksheet objects.

        Args:
            wb (xlsxwriter.Workbook): The Excel workbook object.
            ws (xlsxwriter.Worksheet): The Excel worksheet object.
        """
        self._workbook = wb
        self._worksheet = ws


    def titles_at(self, row:int=0, col:int=0):
        """
        Set the row and column numbers where the titles should be written.

        Args:
            row (int): The row number where the titles should be written.
            col (int): The column number where the titles should be written.
        """
        self._titles_at_row = row
        self._titles_at_col = col


    def after_page_titles(self, plus_rows:int=1) -> int:
        """
        Get the row number after the titles.

        Args:
            plus_rows (int): The number of rows to add after the titles.

        Returns:
            The row number after the titles.
        """
        return len(self.titles) + plus_rows


    def write_page(self):
        """
        Write the Excel page.

        This method writes the titles and tables to the Excel worksheet.
        """
        if not self.titles:
            return
        row = self._titles_at_row
        for title in self.titles:
            self._worksheet.write(
                row, self._titles_at_col, title)
            row += 1
