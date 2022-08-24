import xlsxwriter


class XlsxReport:

    def __init__(self, file_name):
        self.file_name = file_name
        self.pages = []


    def generate(self):
        workbook = xlsxwriter.Workbook(self.file_name)
        self._create_book_pages(workbook)
        workbook.close()


    def _create_book_pages(self, wb):
        """
        Add pages & tables to XlsxReport
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


    def add_book_page(self, name, page):
        self.pages.append(
            {'name': name})
        setattr(self, name, page)






