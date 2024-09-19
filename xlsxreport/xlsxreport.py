import xlsxwriter


class XlsxReport:
    """
    This class is used to create and generate Excel reports.

    Args:
        file_name (str): The name of the Excel file to create.

    Attributes:
        file_name (str): The name of the Excel file to create.
        pages (list): A list of pages to add to the Excel report.

    Methods:
        generate(): Generate the Excel report.
        add_book_page(name, page): Add a page to the Excel report.
    """

    def __init__(self, file_name):
        """
        Create a new XlsxReport object.

        Args:
            file_name (str): The name of the Excel file to create.
        """
        self.file_name = file_name
        self.pages = []


    def generate(self):
        """
        Generate the Excel report.

        This method creates a new Excel workbook and adds all of the pages to the workbook.
        The workbook is then closed.
        """
        workbook = xlsxwriter.Workbook(self.file_name)
        self._create_book_pages(workbook)
        workbook.close()


    def _create_book_pages(self, wb):
        """
        Add pages & tables to XlsxReport

        This method adds all of the pages to the Excel workbook.
        For each page, a new worksheet is created and the page's data is written to the worksheet.
        """
        for pg in self.pages:

            page = getattr(self, pg.get('name'))
            ws = wb.add_worksheet(page.page_name)
            page.set_workbook_worksheet(wb, ws)
            page.write_page()
            for tb in page.tables:
                table = getattr(page, tb.get('name'))
                table.set_workbook_worksheet(wb, ws)
                table.before_write_table()
                table.write_table()
                table.after_write_table()


    def add_book_page(self, name, page):
        """
        Add a page to the Excel report.

        Args:
            name (str): The name of the page to add.
            page (object): The page object to add.
        """
        self.pages.append(
            {'name': name})
        setattr(self, name, page)






